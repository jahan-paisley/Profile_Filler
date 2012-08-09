require 'net/pop'
require 'mail'
require 'debugger'
require 'axlsx'
require 'roo'
require 'tiny_tds'
require 'date'

class Profile_Reader	
	
	def initialize(username, password)
		@username, @password , @nids = username, password, Array.new	
	end

	def email_read
		pop = Net::POP3.new 'mail.tamintelecom.ir'
		pop.enable_ssl(:verify_mode=> OpenSSL::SSL::VERIFY_NONE)
		pop.start @username, @password
		pop.each_mail.reverse_each.each_with_index do |m, i|
			mail = Mail.new m.pop
			puts "processing #{mail.subject}"
			next unless mail.from.to_s.downcase.include? 'bassam' and mail.has_attachments?
			mail.attachments.select {|att| att.filename.end_with?('.xlsx') and (att.filename.downcase.include?("order") or att.filename.downcase.include?("profile"))}.each do |attachment|
				filename = attachment.filename
				puts "\n\twriting attachment: #{filename}"
				begin
					File.open('./inbox/' + filename, "w+b", 0644) {|f| f.write attachment.body.decoded}
					rescue Exception => e
					  puts "Unable to save data for #{filename} because #{e.message}"
				end
		  end
		end
			
	end
	
	def process_file
	  #debugger
		Dir["./inbox/*.xlsx"].each do |file| 
			puts "reading file:", file
			modify_file file
		end
	end
	
	def get_db_info sh_data
		client = TinyTds::Client.new(:username => '', 
									 :password => '', 
									 :host => '')
		#puts sh_data.first[1].to_s
		case  
		  when sh_data.first[1].to_s.start_with?("920")
		    stype= "'Postpaid'"
		  when sh_data.first[1].to_s.start_with?("9218")
		    stype= "'Data'"
		  else
		    stype= "'Prepaid'"
		end 

		query = "WITH T1 AS( 
							SELECT 
								(ROW_NUMBER() OVER(PARTITION BY NationalNo ORDER BY NationalNo ) )AS row,
								registrationlevel, registrationdate, NationalNo,msisdn,firstname,LastName,Nationality, 
								CASE Gender WHEN 'Male' THEN 'مرد' WHEN 'Female' THEN 'زن' ELSE '' END AS Gender, 
								CASE Gender WHEN 'Male' THEN 'آقای' WHEN 'Female' THEN 'خانم' ELSE '' END AS Title, 
								IdentityNo,IssuePlace, CASE MaritalStatus WHEN 'Single' THEN 'مجرد' WHEN 'Married' THEN 'متاهل' ELSE '' END AS MaritalStatus, 
								FatherName, Birthdate, CASE job WHEN 'None' THEN '' WHEN 'Government' THEN 'کارمند بخش دولتی' 
									WHEN 'Nongovernment' THEN 'کارمند بخش خصوصی' WHEN 'Unemployed' THEN 'غیر شاغل/ خانه دار/ بازنشسته' 
									WHEN 'Training' THEN 'دانش آموز/ دانشجو' WHEN 'Industry' THEN 'تکنیسین / کارگر فنی / تعمیرکار' 
									WHEN 'Owner' THEN 'صاحب مغازه' WHEN 'Engineer' THEN 'مهندس/ کارشناس یا مشاور فنی' 
									WHEN 'Health' THEN 'پزشک/ دندانپزشک / داروساز/ دامپزشک / روانپزشک' 
									WHEN 'Lawyer' THEN 'استاد دانشگاه / مشاور / وکیل / قاضی' 
									WHEN 'Manager' THEN 'مدیر عامل/ مدیر ارشد/ مدیر' WHEN 'Workers' THEN 'فروشنده /کارگر ساده' 
									WHEN 'Agriculture' THEN 'کشاورز' WHEN 'Commercial' THEN 'بازرگانی' WHEN 'Service' THEN 'خدمات' 
									WHEN 'Other' THEN 'غیره' ELSE job END AS job , CASE EducationLevel 
									WHEN 'UnderDiploma' THEN 'زیر دیپلم' WHEN 'Diploma' THEN 'دیپلم' WHEN 'College' THEN 'فوق دیپلم' 
									WHEN 'Bachelor' THEN 'لیسانس' WHEN 'Master' THEN 'فوق لیسانس' WHEN 'PhD' THEN 'دکتری' ELSE '' END AS EducationLevel , 
								EMail, PostalCode,'حقیقی' as CustomerType, MobileNo, REPLACE((TelCode+''+ Tel ), '-','') as Tel,
								Province,City, '' ad1,Ave, '' ad2, Street,'' ad3,[Description] , Block, BuildingNo, [Floor],Unit,MunicipalityRegion, 
								ServiceType,DepositAmount,'' Service_List,'' provisiondate,IsCommitted, DepositBank,DepositDate,'' Packages, '' Delivery_Method 
							FROM rightel.dbo.Prospect) 
							
						SELECT  NationalNo,msisdn,firstname,LastName,Nationality,Gender,
								Title,IdentityNo,IssuePlace,MaritalStatus,FatherName,Birthdate,job,EducationLevel,EMail,PostalCode,CustomerType,
								MobileNo,Tel,Province,City,ad1,Ave,ad2,Street,ad3,Description,Block,BuildingNo,Floor,Unit,MunicipalityRegion,ServiceType,
								DepositAmount,Service_List,provisiondate,IsCommitted,DepositBank,DepositDate,Packages,Delivery_Method 
						FROM T1 
						WHERE nationalno IN ( "+sh_data.transpose.first.join(",")+" ) AND registrationlevel =1 AND ServiceType=" + stype +" ORDER BY NationalNo"
		#puts query
    client.execute(query).to_a
	end
	
	def modify_file filename
		excel = Excelx.new filename
		new_excel = Axlsx::Package.new
		shs_data = {}
		
		return if excel.cell(2,'A').downcase.include? 'nationalno' and 
		#debugger
		1.upto(excel.sheets.length) do |i|
			key = excel.sheets[i].to_s
			next if excel.column(1,excel.sheets[i]) == nil or  !(/\d+/ =~ excel.column(1,excel.sheets[i])[2].to_s) 
			nids = excel.column(1,excel.sheets[i])[2..-1]
			msisdns = excel.column(2,excel.sheets[i])[2..-1]
			iccids = excel.column(3,excel.sheets[i])[2..-1]
			shs_data[key] = nids.zip(msisdns,iccids)
		end
		shs_data.each do |sheet_name, sh_data| 
		  next if sh_data.length < 1 
		  result = get_db_info sh_data 
	    new_excel.workbook do |wb|
			  wb.add_worksheet(:name => sheet_name) do |ws|
				  result.each_with_index do |row,i| 			  
					  #debugger
					  if i == 0 then
					    ws.add_row excel.row(1)
					    ws.add_row excel.row(2)
					  end 
					  ws.add_row sh_data.find{|item| item[0] == row.values[0]}.concat(row.values[2..-1])
				  end
			  end
		  end
	  end
		
		new_filename = filename.gsub(/inbox\//,'').gsub(/.xlsx/,"_Filled_#{ DateTime.now.to_s.gsub(/:/,'')}.xlsx")
		puts "writig #{new_filename}"
		new_excel.serialize "#{new_filename}"
	end
end

reader = Profile_Reader.new('','')
reader.email_read
reader.process_file

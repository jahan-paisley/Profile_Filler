require 'net/pop'
require 'mail'
require 'debugger'
require 'axlsx'
require 'roo'
require 'tiny_tds'

class Profile_Reader

	def initialize(username, password)
		@username, @password , @nids = username, password, Array.new	
	end

	def email_read
		pop = Net::POP3.new 'mail.rightel.ir'
		pop.enable_ssl(:verify_mode=> OpenSSL::SSL::VERIFY_NONE)
		pop.start @username, @password
		pop.each_mail.reverse_each.each_with_index do |m, i|
			mail = Mail.new m.pop
			puts "processing #{mail.subject}"
			if mail.has_attachments?
				mail.attachments.each do | attachment |
				 if (attachment.filename.end_with?('.xlsx'))
					filename = attachment.filename
					puts "\twriting attachment: #{filename}"
					begin
					 File.open('inbox/' + filename, "w+b", 0644) {|f| f.write attachment.body.decoded}
					rescue Exception => e
					 puts "Unable to save data for #{filename} because #{e.message}"
					end
				 end
				end
			end
			break if i > 3
		end	
	end
	
	def process_file
		Dir["/Sites/inbox/*.xlsx"].each do |file| 
			puts "reading file:", file
			modify_file file
		end
	end
	
	def modify_file filename
		debugger
		p = Excelx.new filename
		#puts "\t", p.sheets
		p.sheets.each do |sheet|
			@nids << p.column(1)
			break
		end
		
		client = TinyTds::Client.new(:username => '', 
									 :password => '@!', 
									 :host => '2')
		puts @nids.join(",")
		query = "WITH T1 AS( 
							SELECT 
								(ROW_NUMBER() OVER(PARTITION BY NationalNo ORDER BY nationalno ) )AS row,
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
						WHERE nationalno IN ( "+@nids.join(",")+" ) AND registrationlevel =1
						ORDER BY NationalNo "
		puts query
		result = client.execute(query)
		debugger
		result.each { |row| puts row.values } 
		#p.serialize('simple.xlsx')
	end
	
end

reader = Profile_Reader.new('j.','@!')
#reader.email_read
reader.process_file

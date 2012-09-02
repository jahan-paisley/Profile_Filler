# encoding: utf-8
require 'net/imap'
require 'mail'
require 'debugger'
require 'axlsx'
require 'roo'
require 'tiny_tds'
require 'date'
require 'time'
require 'net/smtp'
require 'log4r'
require 'openssl'
require 'zip/zip'

class Profile_Reader
	include Log4r
	def initialize 
		@email_params = {:username=> 'example', :password=> 'example', :host=> 'mail.example.ir', :senders => ['shahmoradi','bassam']}
		@sql_params = {:username=> 'example', :password=> 'example', :host=> '1.1.1.1'}
		outputters = Outputter.stdout
		@path = './inbox/'
		@keywords= ['order', 'profile']
		@files = []
	end
	
	def email_read
	  puts "\nstart searching emails ...\n"
	  senders,@subjects = [], []
	  
	  sstart, eend = set_dates
		puts "start: "+ sstart+ "\tend: "+	 eend
		if @email_params[:senders].length >1 
		  senders << "OR"
		  @email_params[:senders].each{|o| senders << "FROM" << o} 
		else
		  senders = @email_params[:senders]  
		end
		imap = Net::IMAP.new(@email_params[:host],:ssl=>{:verify_mode=> OpenSSL::SSL::VERIFY_NONE}) 
		imap.login(@email_params[:username], @email_params[:password])
		imap.select("INBOX")
		
		#keywords = ["OR","FROM","bassam","FROM","shahmoradi","SINCE", sstart,"BEFORE",eend] 
		keywords = senders + ["SINCE", sstart, "BEFORE",eend]
		msgs = imap.search(keywords)
		puts "found #{msgs.length} messages ..."
		if msgs.length == 0
		  puts "no mail found"
		  exit  
	  end
		 
		msgs.each do |msg_id|
			mail = Mail.new (imap.fetch(msg_id, "RFC822")[0].attr["RFC822"])
			puts "\tProcessing #{mail.subject} ..."
			@subjects << mail.subject
			next unless mail.has_attachments?
			mail.attachments.select{|att| att.filename.end_with?('.xlsx') and @keywords.any?{|key| att.filename.downcase.include?(key)}}.each do |att|
				puts "\t\tdownloading attachment: #{att.filename}"
				begin
					File.open(@path + att.filename, "w+b", 0644) {|f| f.write att.body.decoded}
				rescue Exception=> e
						puts "Unable to save data for #{att.filename} because #{e.message}"
				end
			end
		end
		imap.logout 
		imap.disconnect
		puts "\nDone searching emails"
	end
	
	def set_dates
		format = "%d-%b-%Y"
		last_date = Dir.entries(@path).select{|o| o.end_with?(".xlsx")}.map{|c| File.stat(@path+c).ctime }.max
		start = (last_date || Time.now.to_datetime.prev_day.prev_day)
		sstart = start.strftime(format)
		eend = DateTime.now.next_day.strftime(format)
	  return sstart, eend 
	end
	
	def process_file
	  puts "\nProcessing files ...\n"		
		Dir[@path+"*.xlsx"].each do |file| 
			puts "reading:"+ File.basename(file)
			generate_filled_file (file)
		end
	  puts "\nDone, processing files\n"		
	end
	
	def get_db_info sh_data
		client = TinyTds::Client.new(@sql_params)
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
							
						SELECT	NationalNo,msisdn,firstname,LastName,Nationality,Gender,
								Title,IdentityNo,IssuePlace,MaritalStatus,FatherName,Birthdate,job,EducationLevel,EMail,PostalCode,CustomerType,
								MobileNo,Tel,Province,City,ad1,Ave,ad2,Street,ad3,Description,Block,BuildingNo,Floor,Unit,MunicipalityRegion,ServiceType,
								DepositAmount,Service_List,provisiondate,IsCommitted,DepositBank,DepositDate,Packages,Delivery_Method 
						FROM T1 
						WHERE nationalno IN ( "+sh_data.transpose.first.join(",")+" ) AND registrationlevel =1 AND ServiceType=" + stype +" ORDER BY NationalNo"
		client.execute(query).to_a
	end
	
	def generate_filled_file file_path
		excel, new_excel, shs_data = Excelx.new file_path,Axlsx::Package.new, {}
		return unless excel.cell(2,'A').downcase.include? 'national' 
		
		0.upto(excel.sheets.length-1) do |i|
			excel.default_sheet = excel.sheets[i]	 
			key = excel.sheets[i].to_s
			next if excel.first_row.nil? or !(/\d+/ =~ excel.cell(3,'A').to_s) 
			nids = excel.column(1,excel.sheets[i])[2..-1]
			msisdns = excel.column(2,excel.sheets[i])[2..-1]
			iccids = excel.column(3,excel.sheets[i])[2..-1]
			shs_data[key] = nids.zip(msisdns,iccids)
		end
		excel.default_sheet = excel.sheets[0]
		shs_data.each do |sheet_name, sh_data| 
			next unless sh_data.length > 0 
			result = get_db_info sh_data 
			new_excel.workbook do |wb|
				wb.add_worksheet(:name=> sheet_name) do |ws|
					result.each_with_index do |row,i| 
						if i == 0 then
							ws.add_row excel.row(1)
							ws.add_row excel.row(2)
						end
						begin
						  found = sh_data.find{|item| Integer(item[0]) == Integer(row.values[0])}
						  found[2] += ' '
							ws.add_row found.concat(row.values[2..-1]) 
						rescue
							p "\n\n\n\terror for nationalno: "+row.values[0].to_s() + "in file:"+ file_path
							ws.add_row [ ]
						end
					end
				end
			end
		end
		new_file_path = './output/'+ File.basename(file_path, ".xlsx")+ "_Filled_#{ DateTime.now.to_s.gsub(/:/,'')}.xlsx"
		puts "\twritig #{File.basename(new_file_path)}"
		new_excel.serialize "#{new_file_path}"
		@files << new_file_path
	end
	
	def send_files_to  
	  sleep(2)  
	  make_zip 
	  context = Net::SMTP.default_ssl_context
		context.verify_mode = OpenSSL::SSL::VERIFY_NONE
		smtp = Net::SMTP.new(@email_params[:host], 25)
		smtp.enable_starttls_auto
  	begin
      smtp.start('localhost',@email_params[:username],@email_params[:password], :login) do
  				m = Mail.new
  				m.from = 'j.zinedine@tamintelecom.ir'
  				m.to = 'j.zinedine@tamintelecom.ir'
  				
  				m.body = "This is an automatically generated email."
  				m.subject = @subjects.join(" | ")
  				m.add_file "./files.zip"
  				puts m.from
  				smtp.sendmail m.to_s, 'j.zinedine@tamintelecom.ir','j.zinedine@tamintelecom.ir'
  		end
	  rescue Exception => e
	    p e
	  end
	end
	
	def make_zip
    Zip::ZipOutputStream.open("./files.zip") do |z|
      @files.each do |f|
        title = File.basename(f)
        z.put_next_entry(title)
        z.print IO.read(f)
      end
    end
  end
	
	def test_socket
	   s = timeout(30) { TCPSocket.open('mail.tamintelecom.ir',25) } 
     #logging "TLS connection started"
     #s.sync_close = true
     #s.connect
     #p s.methods
     s.puts 'HELO mail.tamintelecom.ir'
	   puts s.recvmsg
	end
  
end  


reader = Profile_Reader.new
reader.email_read
reader.process_file
reader.send_files_to 
puts "\nDone."


#rPowerPoint
#A plugin for exporting data to pptx format
#by Hao Zheng
#July 2014

require 'fileutils'
require 'csv'
require 'json'


module RPowerPoint
	class PowerPointObject 

		#Slide object corresponding to each slide in template
		class SlideObject
			def initialize
				@hash = {}
				@no_change=false
			end

			def set_no_change
				@no_change=true
			end

			def add_hash(placeholder,substitution)
				@hash[placeholder]= substitution
			end

			def display
				puts @hash
				puts @no_change
			end

			def if_no_change
				return @no_change
			end

			def get_hash
				return @hash
			end

		end


		def initialize(output_name,source_data)




			rename_zip(output_name)
			mac_unzip(output_name)
			@output_file_name=output_name
			create_slides(output_name,source_data)
			delete_folder(output_name)
		end

		#File utilities

		def get_filename
			return @output_file_name+'.pptx'
		end

		def delete_folder(folder_name)
			#Linux
			#FileUtils.rm_rf(folder_name)
			#Mac
			exec('rm -rf '+folder_name)
		end

		def file_write(text,filename)
			if text!=""
				File.open(filename, 'w') { |file| file.write(text) }
			else
				puts 'Error,cannot find placeholder text!'
			end
		end

		#Copy everything under source directory to destnation directory
		def dir_content_copy(source_dir,dest_dir)
			FileUtils.cp_r(Dir[source_dir+'/*'],dest_dir)
		end

		#Make a new directory for output pptx
		def mkdir(output_name)
			FileUtils.mkdir_p output_name
		end

		def rename_zip(output_name)
			exec('mv '+output_name+'/template.pptx '+output_name+'/template.zip')
		end

		#Ditto command only work under Mac, please use approiate linux command

		def mac_compress(output_name)
			exec('ditto -ck --rsrc --sequesterRsrc '+output_name+' '+output_name+'.pptx')
		end

		def mac_unzip(output_name)
			puts 'mv '+output_name+'/template.pptx '+output_name+'/template.zip'
#			exec('ditto -x -k template.zip template')
		end


		#PowerPoint function
		#Copy slides{num}.xml to ppt/slides
		def copy_slides(dest,source,output_name)

			text = File.read('template/ppt/slides/slide'+source.to_s+'.xml')
			file_write(text,output_name+'/ppt/slides/slide'+dest.to_s+'.xml')

			text = File.read('template/ppt/slides/_rels/slide'+source.to_s+'.xml.rels')
			file_write(text,output_name+'/ppt/slides/_rels/slide'+dest.to_s+'.xml.rels')

			return dest		
		end

		#[Content_Types].xml
		def edit_content_type(slide_number,output_name,init_flag)
 
			if init_flag==false
				content_type_xml='<Override PartName="/ppt/slides/slide'+slide_number.to_s+'.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>'
				text = File.read(output_name+'/[Content_Types].xml').gsub('</Types>',content_type_xml)
				file_write(text,output_name+'/[Content_Types].xml')

			else
			#	remove_text='<Override PartName="/ppt/slides/slide'+slide_number.to_s+'.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
			#	text = File.read(output_name+'/[Content_Types].xml').gsub(remove_text,'')				
			end


		end

		#presentation.xml.rels for setup document relationship
		def set_document_rels(slide_number,output_name,init_flag)



			if init_flag==false
				rid=1000+slide_number
				new_rels_xml='<Relationship Id="rId'+rid.to_s+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide'+slide_number.to_s+'.xml"/>
			</Relationships>'
				text = File.read(output_name+'/ppt/_rels/presentation.xml.rels').gsub('</Relationships>',new_rels_xml)
			else
				remove_rels_xml='<Relationship Id="rId'+(slide_number+1).to_s+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide'+slide_number.to_s+'.xml"/>'
				text = File.read(output_name+'/ppt/_rels/presentation.xml.rels').gsub(remove_rels_xml,'')
			end

			file_write(text,output_name+'/ppt/_rels/presentation.xml.rels')

			return rid
		end


		#ppt/presentation.xml
		def add_to_main_presentation(rid,slide_number,output_name,init_flag)

			if init_flag==false
				new_slidID_xml='<p:sldId id="'+(2000+slide_number).to_s + '" r:id="rId'+rid.to_s + '"/> </p:sldIdLst>'
				text = File.read(output_name+'/ppt/presentation.xml').gsub('</p:sldIdLst>',new_slidID_xml)
			else
				slide_numberadd1=slide_number+1
				remove_sliID_xml=/<p:sldId id="(...)" r:id="rId#{slide_numberadd1}/
				remove_sliID_xml=remove_sliID_xml.match(File.read(output_name+'/ppt/presentation.xml')).to_s
				remove_sliID_xml=remove_sliID_xml +'"/>'
				text = File.read(output_name+'/ppt/presentation.xml').gsub(remove_sliID_xml,'')
			end

			file_write(text,output_name+'/ppt/presentation.xml')
		end
 
		#Replace placeholder with targeted text
		def edit_text(placeholder,replacement,dest_slide_number,template_slide_number,output_name,start_flag)
		

			if start_flag == false
				text = File.read('template/ppt/slides/slide'+template_slide_number.to_s+'.xml').gsub(placeholder,replacement)
			else
				text = File.read(output_name+'/ppt/slides/slide'+dest_slide_number.to_s+'.xml').gsub(placeholder,replacement)
			end

			file_write(text,output_name+'/ppt/slides/slide'+dest_slide_number.to_s+'.xml')

		end


		#PowerPoint operations utilities
		def read_mapping(filename)
			slides_objs=[]
   			CSV.foreach(filename) do |row|
		   		slideNum =row[0].to_i
		   		placeholder =row[1]
		   		substitution =row[2]

		   		if slides_objs[slideNum].nil?
		   			slides_objs[slideNum]=SlideObject.new
		   		end

		   		#Set up placeholder:substitution hash
		   		if placeholder.nil?
		   			slides_objs[slideNum].set_no_change
		   		else
		   			slides_objs[slideNum].add_hash(placeholder,substitution)
		   		end
   			end
   			return slides_objs
	   	end

	   	def set_relationship(slide_count,output_name,init_flag)
			edit_content_type(slide_count,output_name,init_flag)
			rid=set_document_rels(slide_count,output_name,init_flag)
			add_to_main_presentation(rid,slide_count,output_name,init_flag)
	   	end

	   	def create_slides(output_name,source_data)
			#Build destnation template
			mkdir(output_name)
			dir_content_copy('template',output_name)
 
			#Read mapping information
			slides_objs=read_mapping('mapping.csv')

			#Start to generate slides
			slide_count=1


			slides_objs.each_with_index do |template_slide,index|

				init_flag=true
				set_relationship(index+1,output_name,init_flag)
				init_flag=false
				if template_slide.if_no_change == true
					copy_slides(slide_count,index+1,output_name)
					set_relationship(slide_count,output_name,init_flag)
					slide_count+=1
				else
					#Edit each slides
					source_data[template_slide.get_hash.first[1]].each_with_index do |src, src_index|
						#For each data table entrey,create a new slides
						copy_slides(slide_count,index+1,output_name)
 
						#Mark for starting modifying each slide xml
						start_flag=false
						#Replace placeholder with targeted text
						template_slide.get_hash.each do |data_pair|
							#puts data_pair[0],source_data[data_pair[1]][src_index],slide_count,index+1
							edit_text(data_pair[0],source_data[data_pair[1]][src_index],slide_count,index+1,output_name,start_flag)
							start_flag=true
						end
						set_relationship(slide_count,output_name,init_flag)
						slide_count+=1
					end
				end
				#Done edit current slide

 
			end
			#Done edit all slides

			#Package it
			mac_compress(output_name)


	   	end




	end
end


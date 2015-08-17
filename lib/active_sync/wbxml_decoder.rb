# - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
# Last update: 	2015-08-12
# Written by: 	Nicolas Marlier 
# Company: 		WePopp
# For documentation, see:
# https://msdn.microsoft.com/en-us/library/dd299442(v=exchg.80).aspx
# http://www.w3.org/1999/06/NOTE-wbxml-19990624/
# 
# Introduction tuto:
# http://blogs.msdn.com/b/openspecification/archive/2013/02/04/how-to-manually-decode-an-activesync-wbxml-stream.aspx
#
#
# - - - - - - - - - - - - - - - - - - - - - - - - - - - -

require 'yaml'
require 'json'
require 'base64'
require 'nokogiri'

module ActiveSync
	module WbxmlDecoder
		def self.compute_wbxml_data wbxml_string
			data = ActiveSync::WbxmlDecoder.string_to_bytes wbxml_string

			# Compute header fields
			version = ActiveSync::WbxmlDecoder.compute_version(data[0])
			public_identifier = ActiveSync::WbxmlDecoder.compute_public_identifier(data[1])
			encoding = ActiveSync::WbxmlDecoder.compute_encoding(data[2])

			# Compute string table
			string_table_length = data[3]
			string_table_raw_data = data[4..(3+string_table_length)]
			string_table_data = ActiveSync::WbxmlDecoder.compute_string_table string_table_raw_data

			# Compute tag table
			tag_table_raw_data = data[(4+string_table_length)..(data.length - 1)]
			tag_table_data = ActiveSync::WbxmlDecoder.compute_tag_table tag_table_raw_data, encoding

			# Build XML hash
			xml_array = ActiveSync::WbxmlDecoder.build_xml_array([], tag_table_data, nil)

			# Returns all that data
			{
				header: {
					version: version,
					public_identifier: public_identifier,
					encoding: encoding,
				},
				string_table: {
					length: string_table_length,
					data: string_table_data
				},
				data: xml_array
			}
		end

		# Convert a string to an array of bytes (8 bits)
		def self.string_to_bytes string
			data = []
			string.each_byte do |i| data << i.to_i end
			data
		end

		# Compute version based on WAP Binary XML Content Format
		# See http://www.w3.org/1999/06/NOTE-wbxml-19990624/, section 'Version Number'
		def self.compute_version version_byte
			"#{(version_byte / 16)+1}.#{(version_byte % 16)}"
		end

		# Compute public identifier based on WAP Binary XML Content Format
		# See http://www.w3.org/1999/06/NOTE-wbxml-19990624/, section 'Document Public Identifier'
		def self.compute_public_identifier public_identifier_byte
			data = {
				0 => "IN STRING TABLE",
				1 => nil,
				2 => "-//WAPFORUM//DTD WML 1.0//EN",
				3 => "-//WAPFORUM//DTD WTA 1.0//EN",
				4 => "-//WAPFORUM//DTD WML 1.1//EN"
			}
			data[public_identifier_byte % 16]
		end

		# Returns an encoding name based on the byte code
		# Reference: http://www.iana.org/assignments/character-sets/character-sets.xml
		def self.compute_encoding encoding_byte
			charsets = YAML.load_file "lib/active_sync/charsets.yml"
			charsets[encoding_byte]
		end

		def self.compute_string_table string_table_raw_data
			string_table_raw_data
		end

		# Computes the tag table
		def self.compute_tag_table tag_table_raw_data, encoding="UTF-8"
			tag_table_data = []
			i = 0
			current_code_page_name = nil
			while i < tag_table_raw_data.length do
				byte = tag_table_raw_data[i]
				# 8-bit is set
				bit8 = byte >= 128
				# 7-bit is set
				bit7 = byte >= 64

				id = byte % 64
				code_tag_name = get_code_tag_name(id, current_code_page_name)
				str = ""
				if code_tag_name == "STR_I"
					i += 1
					while tag_table_raw_data[i] != 0 && i < tag_table_raw_data.length do
						str += tag_table_raw_data[i].chr
						i += 1
					end
					str = str.force_encoding(encoding)
				end

				

				if code_tag_name == "SWITCH_PAGE"
					i += 1
					current_code_page_name = get_code_page_name(tag_table_raw_data[i])
				end

				tag_table_data << {
					byte: byte,
					id: id,
					bit7: bit7,
					bit8: bit8,
					str: str,
					code_page_name: current_code_page_name,
					code_tag_name: code_tag_name
				}
				i+= 1
			end
			tag_table_data
		end

		def self.get_code_tag_name code_tag_byte, code_page_name
			t = Time.now
			@code_tags ||= YAML.load_file("lib/active_sync/code_tags.yml")
			code_page_name ||= "AirSync"

			
			res = if @code_tags['Global'].keys.include? code_tag_byte
				@code_tags['Global'][code_tag_byte]
			elsif @code_tags.keys.include? code_page_name
				@code_tags[code_page_name][code_tag_byte]
			else
				"UNKNOWN"
			end
			res
		end

		def self.get_code_page_name code_page_byte
			@code_pages ||= YAML.load_file("lib/active_sync/code_pages.yml")
			@code_pages[code_page_byte]
		end


		def self.build_xml_array xml_array, data, current_code_page_name

			return xml_array if data.length == 0
			data_byte = data.shift

			code_tag_name = get_code_tag_name(data_byte[:id], data_byte[:code_page_name])
			if code_tag_name == "END"
				return xml_array
			end

			code_page_name_pefix = ""
			if data_byte[:code_page_name] && data_byte[:code_page_name] != "AirSync"
				code_page_name_prefix = "#{data_byte[:code_page_name]}:".downcase
			end

			label = "#{code_page_name_prefix}#{code_tag_name}"
			if code_tag_name == "STR_I"
				data.shift
				return data_byte[:str]
			end

			value = {
				label: label,
				children: []
			}
			new_code_page_name = current_code_page_name
			if current_code_page_name == nil ||
				(data_byte[:code_page_name] && (data_byte[:code_page_name] != current_code_page_name))
				new_code_page_name = data_byte[:code_page_name] || "AirSync"
				value[:attributes] ||= {}
				key = :xmlns
				if data_byte[:code_page_name]
					key = "xmlns:#{new_code_page_name.downcase}".to_sym
				end
				value[:attributes][key] = new_code_page_name
			end


			if code_tag_name == "SWITCH_PAGE"

			else

				if data_byte[:bit7]
					child_xml = build_xml_array([], data, new_code_page_name)
					value[:children] = child_xml
				end
				xml_array << value
			end

			build_xml_array xml_array, data, current_code_page_name
			xml_array
		end



		def self.generate_xml_header header
			attributes = header.select{|k, v| v}.map{|k, v| "#{k}=\"#{v}\""}.join(" ")
			"<?xml #{attributes}?>"
		end

		def self.generate_complete_xml data
			body_xml = generate_xml data[:data]
			#body_xml = decode_timezones body_xml
		<<END
#{generate_xml_header data[:header]}
#{body_xml}
END
		end

		def self.generate_xml array, dec=0
			array.map do |array_element|
				if array_element[:children].class == String
					text = array_element[:children]
				else
					text = "\n#{generate_xml(array_element[:children], dec+1)}\n#{"  " * dec}"
				end

				if (array_element[:attributes] || {}).keys.length == 0
					attributes =  ""
				else
					attributes = " " + (array_element[:attributes] || {}).map{|ka, va| "#{ka}=\"#{va}\""}.join(" ")
				end

				"#{"  " * dec}<#{array_element[:label]}#{attributes}>#{text}</#{array_element[:label]}>"
			end.join("\n")
		end

		def decode_timezone timezone_bytes
		end



		def self.wbxml_to_xml wbxml_string
			return "" if wbxml_string == ""
			data = self.compute_wbxml_data wbxml_string
			generate_complete_xml data
		end


		def self.decode_timezones xml
			nxml = Nokogiri::XML(xml)
			#nxml.remove_namespaces!
			nxml.css("Timezone").each do |timezone_item|
				hash = decode_timezone_string(timezone_item.content)
				timezone_item.content = "ENCODED - #{hash[:standard_name]} (UTC #{hash[:bias]})"
			end
			nxml.to_s
		end

		def self.timezone_hash_to_xml timezone_hash
			timezone_hash.map{|k, v|

			}.join("\n")
		end


		def self.decode_timezone_string encoded_timezone_string
			timezone_string = Base64.decode64 encoded_timezone_string
			timezone_bytes = ActiveSync::WbxmlDecoder.string_to_bytes timezone_string

			bias_bytes = timezone_bytes[0..3]
			standard_name_bytes = timezone_bytes[4..67]
			standard_date_bytes = timezone_bytes[68..83]
			standard_bias_bytes = timezone_bytes[84..87]
			daylight_name_bytes = timezone_bytes[88..151]
			daylight_date_bytes = timezone_bytes[152..167]
			daylight_bias_bytes = timezone_bytes[168..171]

			{
				bias: bytes_array_to_uint32(bias_bytes),

				standard_name: bytes_array_to_unicode_string(standard_name_bytes),
				standard_date: bytes_array_to_time(standard_date_bytes),
				standard_bias: bytes_array_to_uint32(standard_bias_bytes),

				daylight_name: bytes_array_to_unicode_string(daylight_name_bytes),
				daylight_date: bytes_array_to_time(daylight_date_bytes),
				daylight_bias: bytes_array_to_uint32(daylight_bias_bytes)
			}
		end

		def self.bytes_array_to_uint32 bytes_array
			bytes_array.map(&:chr).join.unpack("l")
		end

		def self.bytes_array_to_unicode_string bytes_array
			bytes_array.map(&:chr).join.unpack("S32").select{|i| i>0}.map{|byte| byte.chr}.join
		end

		def self.bytes_array_to_time bytes_array
			pre_data = bytes_array.map(&:chr).join.unpack("S8")
			data = {}
			[:year, :month, :day_of_week, :day_of_month, :hour, :minute, :second, :millisecond].each_with_index do |key, i|
				data[key] = pre_data[i]
			end
			data
		end

	end
end
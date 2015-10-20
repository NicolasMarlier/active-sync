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

module ActiveSync
	module WbxmlEncoder

		def self.encode_hash_to_wbxml hash
			bytes = []

			bytes += generate_header
			bytes += generate_string_table
			bytes += generate_tag_table hash

			bytes.map{|b| b.chr}.join
		end

		def self.generate_string string
			bytes = []
			bytes << 0x03 # String

			string.each_byte do |i|
				bytes << i.to_i
			end

			bytes << 0x00 # End String
			bytes
		end

		def self.generate_opaque_data string
			bytes = []
			bytes << 0xc3 # Opaque data

			string_bytes = []
			string.each_byte do |i|
				string_bytes << i.to_i
			end
			string_bytes_length = string_bytes.length

			#bytes += [128 + string_bytes_length/128, string_bytes_length%128]
      bytes += self.length_to_bytes string_bytes_length
			bytes += string_bytes

			bytes
    end

    # - - - - - - - - - - - - - - - - - - - - -
    # Very strange way to encode integers
    # http://www.w3.org/TR/wbxml/#_Toc443384895
    # - - - - - - - - - - - - - - - - - - - - -
    def self.length_to_bytes length, continuation=false
      if length > 0
        self.length_to_bytes(length / 128, true) + [length % 128 + ((continuation)?128:0)]
      else
        []
      end
    end

		def self.get_code_tag_byte code_names
			code_names_array = code_names.split(":")
			code_page_name = code_names_array[0]
			code_tag_name = code_names_array[1]

			code_tags = YAML.load_file(File.join(__dir__, "code_tags.yml"))
			code_tag_byte = code_tags[code_page_name].select{|byte, name| name == code_tag_name}.keys.first
			raise "No such tag" unless code_tag_byte
			code_tag_byte
		end

		def self.generate_switch code_page_name
			code_pages = YAML.load_file(File.join(__dir__, "code_pages.yml"))
			code_page_byte = code_pages.select{|byte, name| code_page_name == name}.keys.first
			bytes = []
			bytes << 0x00 # Switch
			bytes << code_page_byte
			bytes
		end

		def self.generate_header
			bytes = []
			bytes << 3
			bytes << 1
			bytes << 106
			bytes
		end

		def self.generate_string_table
			bytes = []
			bytes << 0x00
			bytes
		end

		def self.generate_tag_table hash
			bytes = []
			hash.each do |k, v|
				# Switch case
				if k =~ /SWITCH(.*)/
					bytes += generate_switch(v)
				# Case tag containing a hash
				elsif v.class == Hash
					# Case empty hash (got no content)
					if v.keys.length == 0
						bytes << get_code_tag_byte(k)
					# Case non-empty hash (got content)
					else
						bytes << get_code_tag_byte(k) + 0x40
						bytes += generate_tag_table(v)
						bytes << 0x01
					end
				# Case tag containing a string
				elsif v.class == String
					bytes << get_code_tag_byte(k) + 0x40
					if k == "Calendar:Timezone"
						bytes += generate_string(encode_timezone(v))
					elsif k == "ComposeMail:Mime"
						bytes += generate_opaque_data(v)
					else
						bytes += generate_string(v)
					end
					bytes << 0x01
				end
			end
			bytes
		end

		def self.utf_string_to_unicode_string utf_string
			bytes = string_to_bytes(utf_string)
			bytes += (bytes.length..31).map{ 0 }
			bytes.pack("S32")
		end

		# Convert a string to an array of bytes (8 bits)
		def self.string_to_bytes string
			data = []
			string.each_byte do |i| data << i.to_i end
			data
		end

		def self.encode_timezone_date date
			[
			date.year,
			date.month,
			0,
			date.day,
			date.hour,
			date.minute,
			date.second,
			0
			].pack("S8")
		end

		def self.encode_timezone timezone_id

			all_timezones = YAML.load_file(File.join(__dir__, "timezones.yml"))
	    	timezone_hash = build_timezone_hash({}, all_timezones[timezone_id].split("\n"))["VTIMEZONE"]
			
			if timezone_hash['DAYLIGHT']

				standard_bias = bias_from_hour_minute_to_minutes(timezone_hash['STANDARD']['TZOFFSETTO'].to_i)
		    	timezone_hash = {
					bias: -standard_bias,
					standard_name: timezone_hash['STANDARD']['TZNAME'],
					standard_date: DateTime.parse(timezone_hash['STANDARD']['DTSTART']),
					standard_bias: 0,
					daylight_name: timezone_hash['DAYLIGHT']['TZNAME'],
					daylight_date: DateTime.parse(timezone_hash['DAYLIGHT']['DTSTART']),
					daylight_bias: standard_bias - bias_from_hour_minute_to_minutes(timezone_hash['DAYLIGHT']['TZOFFSETTO'].to_i),
				}

				standard_date_string = encode_timezone_date timezone_hash[:standard_date]
				daylight_date_string = encode_timezone_date timezone_hash[:daylight_date]
			else
				timezone_hash = {
					bias: bias_from_hour_minute_to_minutes(timezone_hash['STANDARD']['TZOFFSETTO'].to_i),
					standard_name: timezone_hash['STANDARD']['TZNAME'],
					standard_bias: 0,
					daylight_name: timezone_hash['STANDARD']['TZNAME'],
					daylight_bias: 0,
				}

				standard_date_string = [0,0,0,0,0,0,0,0].pack("S8")
				daylight_date_string = [0,0,0,0,0,0,0,0].pack("S8")
			end

			

			bias_string = [timezone_hash[:bias]].pack("l")
			standard_name_string = utf_string_to_unicode_string timezone_hash[:standard_name]
			standard_bias_string = [timezone_hash[:standard_bias]].pack("l")
			daylight_name_string = utf_string_to_unicode_string timezone_hash[:daylight_name]
			daylight_bias_string = [timezone_hash[:daylight_bias]].pack("l")


			timezone_string = bias_string + standard_name_string + standard_date_string + standard_bias_string + daylight_name_string + daylight_date_string + daylight_bias_string

			Base64.encode64 timezone_string
		end

		def self.bias_from_hour_minute_to_minutes bias
			(bias.abs / 100 * 60 + bias.abs % 100) * ((bias > 0)?1:-1)
		end

	    def self.build_timezone_hash hash, lines
			return hash if lines.length == 0
			line = lines.shift

			line_data = line.split(":")
			key = line_data[0]
			value = line_data[1]

			if key == "BEGIN"
				hash[value] = build_timezone_hash({}, lines)
			elsif key == "END"
				return hash
			else
				hash[key] = value
			end
			build_timezone_hash hash, lines

		end

		def self.example_hash
			{
				"AirSync:Sync" => {
					"AirSync:Collections" => {
						"AirSync:Collection" 	=> {
							"AirSync:SyncKey" 		=> "724540548",
							"AirSync:CollectionId" 	=> "4",
							"AirSync:GetChanges" 	=> {},
							"AirSync:WindowSize" 	=> "25",
							"AirSync:Options" 			=> {
								"AirSync:FilterType" 	=> "5",
								"SWITCH" 				=> "AirSyncBase",
								"AirSyncBase:BodyPreference" 		=> {
									"AirSyncBase:Type" 				=> "1",
									"AirSyncBase:TruncationSize" 	=> "32768"
								}
							},
							"SWITCH" 	=> "AirSync",
							"AirSync:Commands" 	=> {
								"AirSync:Add" 	=> {
									"AirSync:ClientId" 	=> "10473",
									"AirSync:ApplicationData" => {
										"SWITCH" 					=> "Calendar",
										"Calendar:Timezone" 		=> "xP///wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAoAAAAFAAMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMAAAAFAAIAAAAAAAAAxP///w==",
										"Calendar:AllDayEvent" 		=> "0",
										"Calendar:BusyStatus" 		=> "2",
										"Calendar:DtStamp" 			=> "20150811T141751Z",
										"Calendar:EndTime" 			=> "20150811T154500Z",
										"Calendar:Sensitivity" 		=> "0",
										"Calendar:Subject" 			=> "Test event 2",
										"Calendar:StartTime" 		=> "20150811T144500Z",
										"Calendar:UID" 				=> "EB2A3D6274F644879F045A52ECDDA9F01",
										"Calendar:MeetingStatus" 	=> "0"
									}
								}
							}
						}
					}
				}
			}
		end
	end
end
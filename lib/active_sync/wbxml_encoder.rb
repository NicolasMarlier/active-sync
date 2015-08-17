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

			bytes.map(&:chr).join
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

		def self.get_code_tag_byte code_names
			code_names_array = code_names.split(":")
			code_page_name = code_names_array[0]
			code_tag_name = code_names_array[1]

			code_tags = YAML.load_file("lib/active_sync/code_tags.yml")
			code_tag_byte = code_tags[code_page_name].select{|byte, name| name == code_tag_name}.keys.first
			raise "No such tag" unless code_tag_byte
			code_tag_byte
		end

		def self.generate_switch code_page_name
			code_pages = YAML.load_file("lib/active_sync/code_pages.yml")
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
					bytes += generate_string(v)
					bytes << 0x01
				end
			end
			bytes
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
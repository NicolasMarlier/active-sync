# - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
# Last update: 	2015-08-12
# Written by: 	Nicolas Marlier 
# Company: 		WePopp
# 
# Status on sync requests reference:
# https://msdn.microsoft.com/en-us/library/gg675457(v=exchg.80).aspx
# 
# Synckey behaviour:
# https://msdn.microsoft.com/en-us/library/gg663426(v=exchg.80).aspx
# 
# Example in C#
# https://msdn.microsoft.com/en-us/library/office/hh361570(v=exchg.140).aspx
# - - - - - - - - - - - - - - - - - - - - - - - - - - - -
require "httpclient"
require "nokogiri"

module ActiveSync
	class Client
		
		attr_accessor :debug

		def initialize params
			params[:username]
			@client = HTTPClient.new
			@client.set_auth nil, params[:username], params[:password]

			@username 		= params[:username]
			@device_id 		= params[:device_id]
			@device_type	= params[:device_type]
			@policy_key		= params[:policy_key] || 0
			@server_url     = params[:server_url]

			set_policy_key if @policy_key == 0
		end

		def make_request command, hash
			
			url = "#{@server_url}?User=#{@username}&DeviceId=#{@device_id}&DeviceType=#{@device_type}&Cmd=#{command}"

			p url if debug
			if hash != nil
				wbxml_string = WbxmlEncoder.encode_hash_to_wbxml hash
			else
				wbxml_string = "" 
			end

			p "*" * 50 if debug
			p hash if debug
			
			p "*" * 50 if debug
			p wbxml_string if debug

			response = @client.post(url, wbxml_string, {
				'MS-ASProtocolVersion' 	=> '14.1',
				'Content-Type' 			=> "application/vnd.ms-sync.WBXML",
				'X-MS-PolicyKey' 		=> "#{@policy_key}",
				'User-Agent' 			=> "Apple-iPhone7C2/1208.143"
			})

			p "*" * 50 if debug
			print response.header.dump if debug
			print "\n" if debug
			p "*" * 50 if debug

			print response.body if debug
			print "\n" if debug
			p "*" * 50 if @debug

			xml = WbxmlDecoder.wbxml_to_xml response.body
			if debug				
				print xml
				print "\n"
			end

			parse_response xml
		end

		def set_policy_key
			nxml_1 = provision_request_1
			nxml_1.remove_namespaces!
			temporary_policy_key = nxml_1.css("PolicyKey").text

			nxml_2 = provision_request_2 policy_key: temporary_policy_key
			@policy_key = nxml_2.css("PolicyKey").text
		end

		def provision_request_1
			hash = {
			  "SWITCH" => "Provision",
			  "Provision:Provision" => {
			  	"SWITCH_1" => "Settings",
			    "Settings:DeviceInformation" => {
					"Settings:Set" => {
						"Settings:Model" => "Julie",
			            "Settings:IMEI" => "Julie",
			            "Settings:FriendlyName" => "Julie",
			            "Settings:OS" => "Julie",
			            "Settings:OSLanguage" => "en",
			            "Settings:PhoneNumber" => "Julie",
			            "Settings:MobileOperator" => "Julie",
			            "Settings:UserAgent" => "Julie"
					}
			    },
			    "SWITCH_2" => "Provision",
			    "Provision:Policies" => {
			      "Provision:Policy" => {
			        "Provision:PolicyType" => "MS-EAS-Provisioning-WBXML"
			      }
			    }
			  }
			}

			make_request "Provision", hash
		end

		def provision_request_2 params={}
			hash = {
			  "SWITCH" => "Provision",
			  "Provision:Provision" => {
			  	"Provision:Policies" => {
			      "Provision:Policy" => {
			        "Provision:PolicyType" => "MS-EAS-Provisioning-WBXML",
					"Provision:PolicyKey" => "#{params[:policy_key]}",
					"Provision:Status" => "1"
			      }
			    }
			  }
			}

			make_request "Provision", hash
		end

		def folder_sync_request params={}
			params[:sync_key] ||= 0
			hash = {
				"SWITCH" => "FolderHierarchy",
				"FolderHierarchy:FolderSync" => {
					"FolderHierarchy:SyncKey" => "#{params[:sync_key]}"
				}
			}
			make_request "FolderSync", hash
		end

		def sync_request params={}
			params[:sync_key] ||= 0
			raise "Missing param collection_id" unless params[:collection_id]
			hash = {
				"AirSync:Sync" => {
					"AirSync:Collections" => {
						"AirSync:Collection" => {
							"AirSync:SyncKey" => "#{params[:sync_key]}",
							"AirSync:CollectionId" => "#{params[:collection_id]}"
						}
					}
				}
			}
			make_request "Sync", hash
		end

		def create_event_request params={}
			params[:sync_key] ||= 0
			hash = generate_sync_command_hash({
				sync_key: params[:sync_key],
				collection_id: params[:collection_id],
				commands_hash: {
				        "AirSync:Add" => {
				          "AirSync:ClientId" => "#{params[:client_id]}",
				          "AirSync:ApplicationData" => {
				          	"SWITCH_1" => "AirSyncBase",
				          	"AirSyncBase:Body" => {
				          		"AirSyncBase:Type" => "1",
				          		"AirSyncBase:Data" => "#{params[:body]}",
			          		},
				            "SWITCH_2" => "Calendar",
				            "Calendar:Timezone" => "#{params[:timezone]}",
				            "Calendar:AllDayEvent" => "#{params[:all_day_event]}",
				            "Calendar:BusyStatus" => "#{params[:busy_status]}",
				            "Calendar:DtStamp" => "#{params[:dt_stamp]}",
				            "Calendar:EndTime" => "#{params[:end_time]}",
				            "Calendar:Sensitivity" => "#{params[:sensitivity]}",
				            "Calendar:Location" => "#{params[:location]}",
				            "Calendar:Subject" => "#{params[:subject]}",
				            "Calendar:StartTime" => "#{params[:start_time]}",
				            "Calendar:UID" => "#{params[:id]}",
				            "Calendar:MeetingStatus" => "#{params[:meeting_status]}",
				            "Calendar:Attendees" => params[:attendees].map { |attendee|
				            	{
									"Calendar:Attendee" => {
					            		"Calendar:Name" => "#{attendee}",
					            		"Calendar:Email" => "#{attendee}",
					            	}
				            	}
				            }.first
				          }
				        }
				      }
				})
			make_request "Sync", hash
		end

		def delete_event_request params={}
			params[:sync_key] ||= 0
			hash = generate_sync_command_hash({
				sync_key: params[:sync_key],
				collection_id: params[:collection_id],
				commands_hash: {
				        "AirSync:Delete" => {
				          "AirSync:ServerId" => "#{params[:server_id]}"
				        }
				      }
				})
			make_request "Sync", hash
		end

		def fetch_event_request params={}
			params[:sync_key] ||= 0
			hash = generate_sync_command_hash({
				sync_key: params[:sync_key],
				collection_id: params[:collection_id],
				commands_hash: {
				        "AirSync:Fetch" => {
				          "AirSync:ServerId" => "#{params[:server_id]}"
				        }
				      }
				})
			make_request "Sync", hash
		end

		def change_request params={}
			hash = generate_sync_command_hash({
				sync_key: params[:sync_key],
				collection_id: params[:collection_id],
				commands_hash: {
				        "AirSync:Change" => {
				          "AirSync:ServerId" => "#{params[:server_id]}",
				          "AirSync:ApplicationData" => {
				          	"SWITCH_1" => "AirSyncBase",
				          	"AirSyncBase:Body" => {
				          		"AirSyncBase:Type" => "1",
				          		"AirSyncBase:Data" => "#{params[:body]}",
			          		},
				            "SWITCH_2" => "Calendar",
				            "Calendar:Timezone" => "#{params[:timezone]}",
				            "Calendar:AllDayEvent" => "#{params[:all_day_event]}",
				            "Calendar:BusyStatus" => "#{params[:busy_status]}",
				            "Calendar:DtStamp" => "#{params[:dt_stamp]}",
				            "Calendar:EndTime" => "#{params[:end_time]}",
				            "Calendar:Sensitivity" => "#{params[:sensitivity]}",
				            "Calendar:Location" => "#{params[:location]}",
				            "Calendar:Subject" => "#{params[:subject]}",
				            "Calendar:StartTime" => "#{params[:start_time]}",
				            "Calendar:UID" => "#{params[:id]}",
				            "Calendar:MeetingStatus" => "#{params[:meeting_status]}",
				            "Calendar:Attendees" => params[:attendees].map { |attendee|
				            	{
									"Calendar:Attendee" => {
					            		"Calendar:Name" => "#{attendee}",
					            		"Calendar:Email" => "#{attendee}",
					            	}
				            	}
				            }.first
				          }
				        }
				      }
				})
			make_request "Sync", hash
		end

		

		private

		def parse_response xml
			Nokogiri::XML(xml)
		end

		def generate_sync_command_hash params={}
			{
			  "AirSync:Sync" => {
			    "AirSync:Collections" => {
				    "AirSync:Collection" => {
				      "AirSync:SyncKey" => "#{params[:sync_key]}",
				      "AirSync:CollectionId" => "#{params[:collection_id]}",
				      "AirSync:GetChanges" => {},
				      "AirSync:WindowSize" => "25",
				      "AirSync:Options" => {
				        "AirSync:FilterType" => "5",
				        "SWITCH" => "AirSyncBase",
				        "AirSyncBase:BodyPreference" => {
				          "AirSyncBase:Type" => "1",
				          "AirSyncBase:TruncationSize" => "32768"
				        }
				      },
				      "SWITCH" => "AirSync",
				      "AirSync:Commands" => params[:commands_hash] || {}
				    }
				}
			  }
			}
		end

		def generate_invitation_email params={}

			event = {
				:timezone=>{
				:bias=>[-60],
				:standard_name=>"",
				:standard_date=>{:year=>0, :month=>10, :day_of_week=>0, :day_of_month=>5, :hour=>3, :minute=>0, :second=>0, :millisecond=>0},
				:standard_bias=>[0],
				:daylight_name=>"",
				:daylight_date=>{:year=>0, :month=>3, :day_of_week=>0, :day_of_month=>5, :hour=>2, :minute=>0, :second=>0, :millisecond=>0},
				:daylight_bias=>[-60]},
				:start_time=>"20150821T120000Z",
				:end_time=>"20150821T130000Z",
				:summary=>"Hello from Rails",
				:id=>"EacOiPKnAryiwUIhboTXkkhTkjXeGE",
				:server_id=>"4:140"
			}

			attendees = params[:attendees] || []
			start_time = params[:start_time] || DateTime.now
			end_time = params[:start_time] || DateTime.now
			dtstamp = params[:start_time] || DateTime.now

			attendees_data = attendees.map do |attendee|
				"ATTENDEE;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION:MAILTO:#{attendee}"
			end.join("\n")
			ics_data = <<END
BEGIN:VCALENDAR
METHOD:REQUEST
PRODID: JulieDesk
VERSION:2.0
BEGIN:VTIMEZONE
TZID:Pacific Standard Time
BEGIN:STANDARD
DTSTART:20000101T020000
TZOFFSETFROM:-0700
TZOFFSETTO:-0800
RRULE:FREQ=YEARLY;INTERVAL=1;BYDAY=1SU;BYMONTH=11
END:STANDARD
BEGIN:DAYLIGHT
DTSTART:20000101T020000
TZOFFSETFROM:-0800
TZOFFSETTO:-0700
RRULE:FREQ=YEARLY;INTERVAL=1;BYDAY=2SU;BYMONTH=3
END:DAYLIGHT
END:VTIMEZONE
BEGIN:VEVENT
UID:#{params[:event_uid]}
ORGANIZER:MAILTO:#{params[:organizer]}
#{attendees_data}
STATUS:CONFIRMED
X-MICROSOFT-CDO-ALLDAYEVENT:FALSE
BEGIN:VALARM
ACTION:DISPLAY
TRIGGER:-PT15M
END:VALARM
SUMMARY: #{params[:summary]}
LOCATION:
DTSTART;TZID=Pacific Standard Time:#{start_time.strftime("%Y%m%dT%h:%M:%S")}
DTEND;TZID=Pacific Standard Time:#{end_time.strftime("%Y%m%dT%h:%M:%S")}
DTSTAMP: #{dtstamp.strftime("%Y%m%dT%h:%M:%S")}
LAST-MODIFIED: #{DateTime.now.strftime("%Y%m%dT%h:%M:%S")}
CLASS:PUBLIC
END:VEVENT
END:VCALENDAR
END
		mail_string = <<END
MIME-Version: 1.0
Subject: #{params[:summary]}
Thread-Topic: #{params[:summary]}
To: #{(params[:attendees] || []).join(";")}
Content-Type: multipart/alternative;
boundary="---Next Part---"
-----Next Part---
Content-Transfer-Encoding: quoted-printable
Content-Type: text/plain; charset="utf-8"
#{params[:description]}
-----Next Part---
Content-Type: text/calendar; charset="utf-8"; method=REQUEST
Content-Transfer-Encoding: base64
#{Base64.encode64 ics_data}
-----Next Part---
END
		mail_string
		end
	end
end
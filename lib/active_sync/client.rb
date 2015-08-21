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
require "yaml"
require "mail"

module ActiveSync
	class Client
		
		attr_accessor :debug

		def initialize params
			@client = HTTPClient.new
			@client.set_auth nil, params[:username], params[:password]

			@username 		= params[:username]
			@device_id 		= params[:device_id]
			@device_type	= params[:device_type]
			@policy_key		= params[:policy_key]
			@server_url     = params[:server_url]
		end

		def policy_key
			@policy_key ||= generate_policy_key
		end

		def make_request command, hash, params={}
			
			url = "#{@server_url}?User=#{@username}&DeviceId=#{@device_id}&DeviceType=#{@device_type}&Cmd=#{command}"

			if debug
				p url 
			end
			if hash != nil
				wbxml_string = WbxmlEncoder.encode_hash_to_wbxml hash
			else
				wbxml_string = "" 
			end

			if debug
				p "*" * 50
				p hash
				p "*" * 50
				p wbxml_string
				p "*" * 50
				print WbxmlDecoder.wbxml_to_xml wbxml_string
				print "\n"
			end

			response = @client.post(url, wbxml_string, {
				'MS-ASProtocolVersion' 	=> '14.1',
				'Content-Type' 			=> "application/vnd.ms-sync.WBXML",
				'X-MS-PolicyKey' 		=> "#{params[:policy_key] || policy_key}",
				'User-Agent' 			=> "Apple-iPhone7C2/1208.143"
			})

			if debug
				p "*" * 50
				print response.header.dump
				print "\n"
				p "*" * 50

				print response.body
				print "\n"
				p "*" * 50
			end

			xml = WbxmlDecoder.wbxml_to_xml response.body
			if debug				
				print xml
				print "\n"
			end

			parse_response xml
		end

		def generate_policy_key
			nxml_1 = provision_request_1
			nxml_1.remove_namespaces!
			temporary_policy_key = nxml_1.css("PolicyKey").text

			nxml_2 = provision_request_2 policy_key: temporary_policy_key
			nxml_2.css("PolicyKey").text
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

			make_request "Provision", hash, {policy_key: "0"}
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

			make_request "Provision", hash, {policy_key: "0"}
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

		def send_invitation_request params={}
			params[:attendees] ||= []
			boundary = "JulieDesk-" + (0...20).map { (65 + rand(26) + rand(2)*32).chr }.join
			timezones_string = get_timezone_data([params[:start_timezone], params[:end_timezone]].uniq.compact)

			attendees_string = params[:attendees].map do |attendee|
				attendee_string = <<END
ATTENDEE;CN="nmarlier@gmail.com";CUTYPE=INDIVIDUAL;PARTSTAT=NEEDS-ACTION
 ;RSVP=TRUE:mailto:#{attendee}
END
			end.join
			ics_string = <<END
BEGIN:VCALENDAR
CALSCALE:GREGORIAN
METHOD:REQUEST
PRODID:-//Julie Desk//EN
VERSION:2.0
#{timezones_string}BEGIN:VEVENT
#{attendees_string}
CLASS:PUBLIC
CREATED:#{params[:dt_stamp].strftime("%Y%m%dT%H%M%SZ")}
DTEND;TZID=Europe/Paris:#{params[:end_time]}
DTSTAMP:#{params[:dt_stamp].strftime("%Y%m%dT%H%M%SZ")}
DTSTART;TZID=Europe/Paris:#{params[:start_time]}
LAST-MODIFIED:#{params[:dt_stamp].strftime("%Y%m%dT%H%M%SZ")}
ORGANIZER;CN="Nicolas Marlier";EMAIL="#{@username}":m
 ailto:#{@username}
SEQUENCE:0
SUMMARY:#{params[:subject]}
LOCATION:#{params[:location]}
TRANSP:OPAQUE
UID:#{params[:uid]}
X-MICROSOFT-CDO-INTENDEDSTATUS:BUSY
END:VEVENT
END:VCALENDAR
END

			mail_string = <<END
Content-Type: multipart/alternative;
	boundary="#{boundary}"
Content-Transfer-Encoding: 7bit
From: #{@username}
Mime-Version: 1.0 (1.0)
Subject: #{params[:subject]}
Message-Id: <#{params[:client_id]}@marlier.onmicrosoft.com>
Date: #{DateTime.now.strftime("%a, %b %Y %H:%M:%S %Z")}
To: #{params[:attendees].join(";")}


--#{boundary}
Content-Type: text/plain;
	charset=us-ascii
Content-Transfer-Encoding: 7bit

#{params[:body]}

--#{boundary}
Content-Type: text/calendar;
	name=meeting.ics;
	charset=utf-8
Content-Transfer-Encoding: quoted-printable

#{Mail::Encodings::QuotedPrintable.encode(ics_string)}
--#{boundary}--
END
			
			hash = {
				"SWITCH" => "ComposeMail",
				"ComposeMail:SendMail" =>  {
					"ComposeMail:ClientId" => "#{params[:client_id]}",
					"ComposeMail:SaveInSentItems" => {},
					"ComposeMail:Mime" => mail_string
				}
			}

			make_request "SendMail", hash
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

		def update_event_request params={}
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

		def get_timezone_data timezones
	      all_timezones = YAML.load_file("lib/active_sync/timezones.yml")
	      timezones.uniq.map{|timezone_id|
	        timezone_data = all_timezones[timezone_id]
	        raise "Unknown timezone: '#{timezone_id}'" unless timezone_data
	        timezone_data
	      }.join("\n")
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
PRODID:-//JulieDesk//EN
VERSION:2.0
CALSCALE:GREGORIAN
#{timezone}
BEGIN:VEVENT
UID:#{params[:event_uid]}
ORGANIZER:MAILTO:#{params[:organizer]}
#{attendees_data}
STATUS:CONFIRMED
SUMMARY:#{params[:summary]}
LOCATION:#{params[:location]}
DTSTART;TZID=Pacific Standard Time:#{start_time.strftime("%Y%m%dT%h:%M:%S")}
DTEND;TZID=Pacific Standard Time:#{end_time.strftime("%Y%m%dT%h:%M:%S")}
CREATED:#{dtstamp.strftime("%Y%m%dT%h:%M:%S")}
END:VEVENT
END:VCALENDAR
END
	
		mail = Mail.new do
	      from    		"julie@julidesk.com"
	      to      		"nmarlier@gmail.com"
	      subject 		"#{params[:summary]}"
	      content_type 	"text/calendar; charset=UTF-8; method=request"
	      body 			"#{params[:description]}"
	    end
	    mail.attachments['invite.ics'] << ics_data
	    mail.attachments['invite.ics'].content_type += "; method=request"
		
		mail.to_s
		end
	end
end
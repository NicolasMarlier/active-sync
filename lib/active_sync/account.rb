require "base64"
require "yaml"

module ActiveSync
	class Account

		attr_accessor :username, :password, :device_id, :device_type, :client, :server_url

		def sync_keys
			@sync_keys ||= {}
		end

		def get_sync_key collection_id
			sync_keys[collection_id] ||= "0"
		end

		def set_sync_key collection_id, sync_key
			sync_keys[collection_id] = sync_key
		end

		def initialize params
			self.username = params[:username]
			self.password = params[:password]
			self.device_id = params[:device_id]
			self.device_type = params[:device_type]
			self.server_url = params[:server_url]
			

			self.client = ActiveSync::Client.new({
				username: self.username,
				password: self.password,
				device_id: self.device_id,
				device_type: self.device_type,
				server_url: self.server_url
				})
		end

		def self.example
			secret_data = YAML.load_file("secret.yml")
			self.new({
				username: secret_data['example_username'],
				password: secret_data['example_password'],
				server_url: secret_data['example_server_url'],
				device_id: "JULIEDESK",
				device_type: "Julie"
				})
		end

		def list_calendars
			response = client.folder_sync_request sync_key: 0
			response.xpath("//folderhierarchy:Add").select do |folder_item|
				[8, 13].include? folder_item.xpath("folderhierarchy:Type").text.to_i
			end.map do |folder_item|
				{
					calendar_id: folder_item.xpath("folderhierarchy:ServerId").text.to_i,
					name: folder_item.xpath("folderhierarchy:DisplayName").text,
					main: folder_item.xpath("folderhierarchy:Type").text.to_i == 8
				}
			end
		end

		def list_events params={}
			raise "Missing param calendar_id" unless params[:calendar_id]
			new_events = []
			deleted_events = []
			updated_events = []

			if get_sync_key(params[:calendar_id]) == "0"
				response = make_request(:sync_request, {
					sync_key: get_sync_key(params[:calendar_id]),
					collection_id: params[:calendar_id]
				})
			end

			continue = true
			while continue do
				response = make_request(:sync_request, {
					sync_key: get_sync_key(params[:calendar_id]),
					collection_id: params[:calendar_id]
				})
				response.remove_namespaces!
				response.css("Add").each do |event_item|
					new_events << parse_event(event_item)
				end
				response.css("Change").each do |event_item|
					updated_events << parse_event(event_item)
				end
				response.css("Delete").each do |event_item|
					deleted_events << parse_deleted_event(event_item)
				end
				continue = response.css("MoreAvailable").length > 0
			end
			{
				new_events: new_events,
				updated_events: updated_events,
				deleted_events: deleted_events
			}
		end



		def create_event params={}
			
			[:calendar_id, :summary, :start_time, :end_time].each do |key|
				raise "Missing param #{key}" unless params[key]
			end
			@sync_keys = {}
			list_events calendar_id: params[:calendar_id]
			uid = generate_uid
			client.debug = true
			response = make_request(:create_event_request, {
				sync_key: get_sync_key(params[:calendar_id]),
				collection_id: params[:calendar_id],
				client_id: 1,
				timezone: "xP///wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAoAAAAFAAMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMAAAAFAAIAAAAAAAAAxP///w==",
				all_day_event: 0,
				busy_status: 2,
				dt_stamp: DateTime.now.strftime("%Y%m%dT%H%M%SZ"),
				start_time: params[:start_time].strftime("%Y%m%dT%H%M%SZ"),
				end_time: params[:end_time].strftime("%Y%m%dT%H%M%SZ"),
				sensitivity: 0,
				location: params[:location],
				subject: params[:summary],
				body: params[:description],
				id: uid,
				meeting_status: 0,
				attendees: params[:attendees]
			})

			{
				event_id: response.css("Add ServerId").text,
				calendar_id: params[:calendar_id],
				uid: generate_uid
			}
		end

		def delete_event params={}
			@sync_keys = {}
			list_events calendar_id: params[:calendar_id]

			make_request(:delete_event_request, {
					sync_key: get_sync_key(params[:calendar_id]),
					collection_id: params[:calendar_id],
					server_id: params[:event_id]
				})
		end

		def update_event params={}
			@sync_keys = {}
			events = list_events calendar_id: params[:calendar_id]

			event = events[:new_events].select{|event| event[:server_id] == params[:event_id]}.first

			start_time = params[:start_time] || DateTime.parse(event[:start_time])
			end_time = params[:end_time] || DateTime.parse(event[:end_time])
			summary = params[:summary] || event[:summary]
			
			make_request(:change_request, {
					sync_key: get_sync_key(params[:calendar_id]),
					server_id: params[:event_id],
					collection_id: params[:calendar_id],
					client_id: 1,
					timezone: "xP///wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAoAAAAFAAMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMAAAAFAAIAAAAAAAAAxP///w==",
					all_day_event: 0,
					busy_status: 2,
					dt_stamp: DateTime.now.strftime("%Y%m%dT%H%M%SZ"),
					start_time: start_time.strftime("%Y%m%dT%H%M%SZ"),
					end_time: end_time.strftime("%Y%m%dT%H%M%SZ"),
					sensitivity: 0,
					location: params[:location],
					subject: summary,
					body: params[:description],
					id: event[:id],
					meeting_status: 0,
					attendees: params[:attendees]
				})
		end

		

		private

		def make_request request, params
			response = client.send(request, params)
			sync_key = response.css("SyncKey").text
			if sync_key.length > 0
				set_sync_key params[:collection_id], sync_key
			end
			
			p "*" * 50
			p "New sync key: #{get_sync_key(params[:collection_id])}"
			p "*" * 50
			response
		end

		def generate_uid
			(0...30).map { (65 + rand(26) + rand(2)*32).chr }.join
		end

		def parse_event event_item
			print event_item.to_s
			{
				timezone: ActiveSync::WbxmlDecoder.decode_timezone_string(event_item.css("Timezone").text),
				start_time: event_item.css("StartTime").text,
				end_time: event_item.css("EndTime").text,
				summary: event_item.css("Subject").text,
				id: event_item.css("UID").text,
				server_id: event_item.css("ServerId").text
			}
		end

		def parse_deleted_event event_item
			{
				server_id: event_item.css("ServerId").text
			}
		end

	end
end
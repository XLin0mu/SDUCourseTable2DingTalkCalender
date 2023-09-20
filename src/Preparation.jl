using XLSX
using Dates
using TimeZones
using DataFrames

function split_course_item(item::String)
    item[begin] == '\n' ? item = chop(item, head = 1, tail=0) : nothing
    item[end] == '\n' ? item = chop(item) : nothing
    return String.(split(item, '\n'))
end

"""
return (dayOfWeek, (begin_time, end_time))"""
function get_grid_period(grid_period::Tuple{Int64, Int64};
    week_begin::Int64 = Monday,
    course_period::Vector{Tuple{Time, Time}} = [
        (Time(8), Time(9, 50)),
        (Time(10, 10), Time(12)),
        (Time(14), Time(15, 50)),
        (Time(16, 10), Time(18)),
        (Time(19), Time(20, 50))
    ])
    return (Day(week_begin + grid_period[2] -1), course_period[grid_period[1]])
end

function get_course_items(xlsxfile::String; work_table::String = "Sheet1", index_ref::String = "B4:H8")
    table = XLSX.readdata(xlsxfile, work_table, index_ref)
    #items = Tuple{String, Tuple{Int64, Int64}}[]
    items = Tuple{Vector{String}, Tuple{Day, Tuple{Time, Time}}}[]
    for row in 1 : size(table)[1], col in 1 : size(table)[2]
        splits = split(table[row, col], "\n\n")
        if length(splits) == 1
            push!(items, (split_course_item(String(table[row, col])), get_grid_period((row, col))))
        else
            for i in eachindex(splits)
                if splits[i] != "" && length(split_course_item(String(splits[i]))) < 6
                    if length(splits) == i
                        @error "The course grid at \"row $(row) col $(col)\" missing something, please check it in .xlsx file."
                    end
                    splits[i] *= splits[i+1]
                    splits[i+1] = ""
                    @warn "The course grid at row \"$(row) col $(col)\" may missing something, please using function \"check_course_table\" for checking."
                    break
                end
                push!(items, (split_course_item(String(splits[i])), get_grid_period((row, col))))
            end
        end
    end
    return items
end

function get_course_week_num(course_period_item::String)
    weeks = Any[collect(eachmatch(r"(?<=第).*?(?=周)", course_period_item))...]
    for i in eachindex(weeks)
        ws = parse.(Int, split(weeks[i].match, '-'))
        length(ws) == 1 ? weeks[i] = Week(ws[1]) : weeks[i] = (Week(ws[1]), Week(ws[2]))
    end
    return weeks
end

function new_event()
    return Dict(
        #日程标题
        "summary" => "",
        #日程描述
        "description" => "",
        #日程开始时间
        "start" => Dict{String, Any}(
            #"2023-01-01T00:00:00+08:00"
            "dateTime" => "",
            #Asia/Shanghai
            "timeZone" => ""
        ),
        #日程结束时间
        "end" => Dict{String, Any}(
            #"2023-01-01T00:00:00+08:00"
            "dateTime" => "",
            #Asia/Shanghai
            "timeZone" => ""
        ),
        #是否全天日程
        "isAllDay" => false,
        #日程参与人列表
        "attendees" => [Dict(
            "id" => "",
            "isOptional" => true
        )],
        #日程地点
        "location" => Dict(
            "displayName" => ""
        ),
        #日程提醒
        "reminders" => [Dict(
            "method" => "dingtalk",
            "minutes" => 30
        )],
        #SON格式的扩展能力开关
        "extra" => Dict(
            "noChatNotification" => "true",
            "noPushNotification" => "true"
        ),
    )
end

function modidy_course_event_datetime!(event, date, times; time_zone = tz"Asia/Shanghai")
    event["start"]["dateTime"] = Dates.format(ZonedDateTime(DateTime(date, times[1]), time_zone), "yyyy-mm-ddTHH:MM:SSzzzz")
    event["start"]["timeZone"] = time_zone.name
    event["end"]["dateTime"] = Dates.format(ZonedDateTime(DateTime(date, times[2]), time_zone), "yyyy-mm-ddTHH:MM:SSzzzz")
    event["end"]["timeZone"] = time_zone.name
    return nothing
end

"""
Such as: course_start_date = Date(2023,9,4)"""
function deal_course_table(course_items::Vector{Tuple{Vector{String}, Tuple{Day, Tuple{Time, Time}}}}, course_start_date::Date)
    events = Dict[]
    for item in course_items
        if item[1] != [" "]
            week_nums = get_course_week_num(String(item[1][5]))
            for week_num in week_nums
                event = new_event()
                modidy_course_event!(event, item[1])
                if typeof(week_num) == Week
                    date = course_start_date + week_num - Week(1) + item[2][1] - Day(1)
                    modidy_course_event_datetime!(event, date, item[2][2])
                    push!(events, event)
                else
                    date = course_start_date + week_num[1] - Week(1) + item[2][1] - Day(1)
                    recurrence_end_date = course_start_date + week_num[2] - Week(1) + item[2][1] - Day(1)
                    modidy_course_event_datetime!(event, date, item[2][2])

                    recurrence = Dict(
                        "pattern" => Dict(
                            "type" => "weekly",
                            "daysOfWeek" => lowercase(dayname(dayofweek(date))),
                            "interval" => 1
                            ),
                        "range" => Dict(
                            "type" => "endDate",
                            "endDate" => Dates.format(ZonedDateTime((DateTime(recurrence_end_date, item[2][2][2])), tz"Asia/Shanghai"), "yyyy-mm-ddTHH:MM:SSzzzz"),
                      )
                    )
                    push!(event, "recurrence" => recurrence)

                    push!(events, event)
                end
            end
        end
    end
    return events
end

function modidy_course_event!(event, course_item::Vector{String})
    length(course_item) != 6 ? error("invalid course item as follow:...\n\n$(course_item)\n\n(as below)") : nothing
    event["summary"] = course_item[1]
    event["description"] = "课程名："*course_item[1]*"\n教师："*course_item[4]*"\n上课地点："*course_item[6]
    event["location"] = course_item[6]
    return nothing
end

function check_course_repeat!(events::Vector{Dict})
    for i_event in eachindex(events)
        for i_other_event in eachindex(events)
            events[i_event] == events[i_other_event] && i_event != i_other_event ? events[i_other_event] = Dict() : nothing
        end
    end
end

function modidy_event_attendees!(event::Dict, attendees_id::String)
    event["attendees"][1]["id"] = attendees_id
    return nothing
end
function modidy_event_attendees!(events::Vector{Dict}, attendees_id::String)
    for event in events
        modidy_event_attendees!(event, attendees_id)
    end
    return nothing
end

function generate_course_events(xlsxfile, course_start_date, attendees_id)
    course_items = get_course_items(xlsxfile);
    events = deal_course_table(course_items, course_start_date);
    check_course_repeat!(events);
    modidy_event_attendees!(events, attendees_id);
    return events
end

function check_course_table(events)

end
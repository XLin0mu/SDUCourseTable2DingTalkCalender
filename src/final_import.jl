





include("./Preparation.jl")
include("./DingTalkCalendarLib.jl")
using Main.DingTalkCalendarLib
using Main.Preparation


using Dates
include("../data/config.jl")
#include("../data/config_test.jl")


access_token = get_token(AppKey, AppSecret)
userId = get_userId(phoneNumber, access_token)
unionId = get_unionId(userId, access_token)

#standard
str = """
复变函数与积分变换
2024-06-28
14:30
16:30
兴隆山群楼E座5202d
15
"""

function exam_import(str::String; unionId=unionId)
    ar = split(str, '\n')
    ar[1] = "[考试]" * ar[1]
    ar[6] = "座位号 " * ar[6]

    ne = new_event()
    ne["summary"] = ar[1]
    ne["description"] = ar[6]
    ne["start"] = Dict{String,Any}(
        "dateTime" => ar[2] * 'T' * ar[3] * ":00+08:00",
        "timeZone" => "Asia/Shanghai"
    )
    ne["end"] = Dict{String,Any}(
        "dateTime" => ar[2] * 'T' * ar[4] * ":00+08:00",
        "timeZone" => "Asia/Shanghai"
    )
    ne["location"] = ar[5]
    ne["reminders"] = [Dict(
            "method" => "dingtalk",
            "minutes" => 15
        ), Dict(
            "method" => "dingtalk",
            "minutes" => 30
        ), Dict(
            "method" => "dingtalk",
            "minutes" => 60
        ), Dict(
            "method" => "dingtalk",
            "minutes" => 1440
        )]
    ne["attendees"] = [Dict(
        "id" => unionId,
        "isOptional" => false
    )]

    calendar_create_events(CalendarEvents(ne), access_token, unionId)

end

exam_import(str)



"""


"""



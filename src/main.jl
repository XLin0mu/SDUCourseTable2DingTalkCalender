
#cd("E:\\AhMyGit\\2_MyGitHub\\Mine\\SDUCourse2DingTalkCalender\\src")
#access_token = 1
#unionId = 1
#debug args is upon
include("./DingTalkCalendarAPI/DTCalendarAPI.jl")
include("./DingTalkCalendarAPI/lib.jl")
include("./xlsx2Dict.jl")

function runme(xlsx::String, access_token, unionId)
    lessons_set = xlsx2DingTalkDict(xlsx)
    for lesson in lessons_set
        calendar_create_events(
            CalendarEvents(lesson), access_token, unionId
        )
    end
end

runme(table_dir, access_token, unionId)
using Pkg
Pkg.add(["HTTP", "TOML", "JSON", "XLSX", "Dates", "TimeZones", "DataFrames"])

include("./src/Preparation.jl");
include("./src/DingTalkCalendarLib.jl");

using Dates
include("./data/config.jl");
#include("./data/config_test.jl");

using Main.DingTalkCalendarLib

xlsxfile = "data/table.xlsx"

events = generate_course_events(xlsxfile, course_start_date, attendees_id)
access_token = get_token(AppKey, AppSecret)
userId = get_userId(phoneNumber, access_token)
unionId = get_unionId(userId, access_token)

for event in CalendarEvents.(events)
    calendar_create_events(event, access_token, unionId)
end
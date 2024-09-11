using Pkg
Pkg.add(["HTTP", "TOML", "JSON", "XLSX", "Dates", "TimeZones", "DataFrames"])

include("./Preparation.jl");
include("./DingTalkCalendarLib.jl");

using Dates
include("../data/config.jl");
#include("../data/config_test.jl");

using Main.DingTalkCalendarLib
using Main.Preparation

xlsxfile = "../data/table.xlsx"

events = generate_course_events(xlsxfile, course_start_date, attendees_ids)
userInfo = get_user_info(AppKey, AppSecret, phoneNumber)

for event in events
    calendar_create_events(event, userInfo.access_token, userInfo.unionId)
end
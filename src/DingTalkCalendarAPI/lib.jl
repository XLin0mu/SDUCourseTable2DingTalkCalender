using HTTP
using TOML
using JSON

"""
获取access_token
"""
function get_token(appKey::String, appSecret::String)
    url = "https://api.dingtalk.com/v1.0/oauth2/accessToken"
    header = Dict(
        "Content-Type" => "application/json"
    )
    body = Dict(
        "appKey" => appKey,
        "appSecret" => appSecret
    )
    resp = HTTP.post(url, header, JSON.json(body))
    return JSON.parse(String(resp.body))["accessToken"]
end

"""
通过手机号获取userId
"""
function get_userId(phoneNumber::String, accessToken::String)
    url = "https://oapi.dingtalk.com/topapi/v2/user/getbymobile?access_token="*accessToken
    header = Dict(
        "Content-Type" => "application/json"
    )
    body = Dict(
        "mobile" => phoneNumber
    )
    resp = HTTP.post(url, header, JSON.json(body))
    return JSON.parse(String(resp.body))["result"]["userid"]
end

"""
获取unionId
"""
function get_unionId(userId::String, accessToken::String)
    url = "https://oapi.dingtalk.com/topapi/v2/user/get?access_token="*accessToken
    header = Dict(
        "Content-Type" => "application/json"
    )
    body = Dict(
        "userid" => userId
    )
    resp = HTTP.post(url, header, JSON.json(body))
    return JSON.parse(String(resp.body))["result"]["unionid"]
end

"""
marker for Dict
"""
struct CalendarEvents
    config::Dict
end

"""
创建日历日程
"""
function calendar_create_events(
    event::CalendarEvents,
    accessToken::String,
    userId::String;
    calendarId::String = "primary")

    url = "https://api.dingtalk.com/v1.0/calendar/users/"*userId*"/calendars/"*calendarId*"/events"
    header = Dict(
        "Content-Type" => "application/json",
        "x-acs-dingtalk-access-token" => accessToken
    )
    body = event.config
    resp = HTTP.post(url, header, JSON.json(body))
    println("Adding Schedule $(event.config["summary"]) Succeed!")
end

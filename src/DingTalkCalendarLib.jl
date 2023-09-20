module DingTalkCalendarLib
export get_token, get_userId, get_unionId, get_user_info, CalendarEvents, calendar_create_events
using HTTP
using TOML
using JSON

"""
获取access_token"""
function get_token(app_key::String, app_secret::String)
    url = "https://api.dingtalk.com/v1.0/oauth2/accessToken"
    header = Dict(
        "Content-Type" => "application/json"
    )
    body = Dict(
        "appKey" => app_key,
        "appSecret" => app_secret
    )
    resp = HTTP.post(url, header, JSON.json(body))
    return JSON.parse(String(resp.body))["accessToken"]
end

"""
通过手机号获取userId"""
function get_userId(phone_number::String, access_token::String)
    url = "https://oapi.dingtalk.com/topapi/v2/user/getbymobile?access_token="*access_token
    header = Dict(
        "Content-Type" => "application/json"
    )
    body = Dict(
        "mobile" => phone_number
    )
    resp = HTTP.post(url, header, JSON.json(body))
    return JSON.parse(String(resp.body))["result"]["userid"]
end

"""
获取unionId"""
function get_unionId(user_id::String, access_token::String)
    url = "https://oapi.dingtalk.com/topapi/v2/user/get?access_token="*access_token
    header = Dict(
        "Content-Type" => "application/json"
    )
    body = Dict(
        "userid" => user_id
    )
    resp = HTTP.post(url, header, JSON.json(body))
    return JSON.parse(String(resp.body))["result"]["unionid"]
end

function get_user_info(app_key::String, app_secret::String, phone_number::String)
    #get access token
    access_token = get_token(app_key, app_secret)

    #get userId
    userId = get_userId(phone_number, access_token)

    #get unionId
    unionId = get_unionId(userId, access_token)

    return (access_token = access_token, userId = userId, unionId = unionId)
end

"""
数据类型标记"""
struct CalendarEvents
    config::Dict
end

"""
创建日历日程"""
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
end
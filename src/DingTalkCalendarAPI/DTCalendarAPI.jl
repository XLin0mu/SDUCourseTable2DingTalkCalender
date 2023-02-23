include("./lib.jl")
config = TOML.parsefile("../config.toml")

#get access token
access_token = get_token(
    config["AppKey"],
    config["AppSecret"]
)

#get userId
userId = get_userId(config["phoneNumber"], access_token)

#get unionId
unionId = get_unionId(userId, access_token)
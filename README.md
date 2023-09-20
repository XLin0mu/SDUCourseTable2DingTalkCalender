扩展性挺高，在 `src/xlsx2Dict.jl` 里面有注释（照着默认变量改就行，别碰函数）有外校人想用的话麻烦发个issue，我想后续可以把学校做成参数，这样以后别人调就不用当调参侠了。（没人用就懒得写了，破包快把我搞疯了）

需要自己把.xls文件用excel打开，然后另存为.xlsx格式

钉钉信息写到 info.toml 里面

运行 `./src/main.jl` 后，调整参数 `table_dir` 为课程表的路径名即可。

# 注意

非山大首次使用需要自行修改 dictConfig.jl 中内容。

# info.toml

钉钉企业内部应用开发的参数，还有改权限，哪里卡住发issue

首先注册企业，流程很简单。

成为开发者：

https://open.dingtalk.com/document/orgapp/become-a-dingtalk-developer

创建应用并获取 `AppKey` 和 `AppSecret`：

https://open.dingtalk.com/document/orgapp/create-orgapp

权限：

https://open.dingtalk.com/document/orgapp/permission-management

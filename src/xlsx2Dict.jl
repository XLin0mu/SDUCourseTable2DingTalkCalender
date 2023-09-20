module xlsx2Dict
export xlsx2DingTalkDict

using XLSX
using Dates
using TimeZones

function deal_discrete_gap_periods!(periods, scatter_start, scatter_end, lessons, lesson_struct, term_begin_at)
    for i in scatter_start : scatter_end
        lesson = deepcopy(lesson_struct)
        lesson["start"]["dateTime"] = term_begin_at[1] + Week(periods[i][1]-1)
        lesson["end"]["dateTime"] = term_begin_at[1] + Week(periods[i][1]-1)
        pop!(lesson, "recurrence")
        push!(lessons, deepcopy(lesson))
    end
end

function deal_gap_periods!(periods, gap_start, gap_end, lessons, lesson_struct, term_begin_at)
    #discrete situation
    if gap_end - gap_start < 3
        deal_discrete_gap_periods!(periods, gap_start, gap_end, lessons, lesson_struct, term_begin_at)
        return nothing
    else
        #estimate discrete or regular situation
        bool_vec = [periods[i+2][1] + periods[i][1] == 2periods[i+1][1]
            for i in gap_start : gap_end - 2]

        #initialize check point
        chk_at = false
        chk_point = 1
        #traverse all bool_vec
        for index_bool in eachindex(bool_vec)
            if (bool_vec[index_bool] != chk_at ? (index_bool -= 1; true) : false) || index_bool == gap_end - gap_start -2 +1    #是否进行结算
                if !chk_at
                    #settle down befores
                    scatter_start = gap_start-1 + chk_point
                    scatter_end = gap_start-1 + index_bool
                    if scatter_start <= scatter_end
                        deal_discrete_gap_periods!(periods, scatter_start, scatter_end, lessons, lesson_struct, term_begin_at)
                    end
                else
                    #settle down befores
                    lesson = deepcopy(lesson_struct)
                    gap_of_week = periods[index_bool+1][1] - periods[index_bool][1]
                    lesson["start"]["dateTime"] = lesson["end"]["dateTime"] = term_begin_at[1] + Week(periods[gap_start-1+chk_point][1]-1)
                    lesson["recurrence"]["pattern"]["interval"] = gap_of_week
                    lesson["recurrence"]["range"] = Dict(
                        "type" => "numbered",
                        "numberOfOccurrences" => index_bool+2 - chk_point + 1   #钉钉日历中，重复n次实际上会包含本次
                    )
                    push!(lessons, deepcopy(lesson))
                end
                index_bool += 1
                #modify check point
                chk_at = !chk_at
                chk_point = index_bool
            end
        end
    end
end

function deal_lesson_periods!(lessons, lesson_struct, periods, term_begin_at)
    if length(periods) != 0
        l = 1
        gap_start = 1
        gap_end = 0
        lesson = deepcopy(lesson_struct)
        while l <= length(periods)
            if length(periods[l]) == 1
                gap_end = l
                l += 1
            elseif length(periods[l]) == 2

                #dispose lesson
                lesson["start"]["dateTime"] = lesson["end"]["dateTime"] = term_begin_at[1] + Week(periods[l][1]-1)
                lesson["recurrence"]["range"]["endDate"] =
                Dates.format(ZonedDateTime((term_begin_at[1] + Week(periods[l][2])), term_begin_at[2]), "yyyy-mm-ddTHH:MM:SSzzzz")
                push!(lessons, deepcopy(lesson))
                lesson = deepcopy(lesson_struct)
                l+=1

                #settle gap periods before
                if gap_end >= gap_start
                    deal_gap_periods!(periods, gap_start, gap_end, lessons, lesson_struct, term_begin_at)
                end
                #count from next one
                gap_start = l
                gap_end = l - 1
            else
                throw(ArgumentError("""wrong args in ruleofTable["course_struct"]["course_periods"]"""))
            end
        end

        #settle remained gap periods
        if gap_end >= gap_start
            deal_gap_periods!(periods, gap_start, gap_end, lessons, lesson_struct, term_begin_at)
        end
    end
end

function deal_course_cell(default_struct::Dict{String, Any}, course_cell::String, term_begin_at::Tuple{Date, VariableTimeZone}, ruleofCourseCell::Dict{String, Any})
    #initialize lessons
    lessons = Set{Dict{String, Any}}()
    lesson = deepcopy(default_struct)
    periods = Vector{Int}[]

    #sperate cell into fields
    cell = split(course_cell, ruleofCourseCell["element_split"])

    #traverse every field, edit dict's value, except time
    for i in eachindex(cell)

        if ruleofCourseCell["cell_struct"][i] !== nothing

            if ruleofCourseCell["cell_struct"][i] == "course name"
                lesson["summary"] = ruleofCourseCell["syntax"]["course name"](String(cell[i]))

            elseif ruleofCourseCell["cell_struct"][i] == "description"
                lesson["description"] = ruleofCourseCell["syntax"]["description"](String(cell[i]))

            elseif ruleofCourseCell["cell_struct"][i] == "location"
                lesson["location"]["displayName"] = ruleofCourseCell["syntax"]["location"](String(cell[i]))

            #construct time period of lessons
            elseif ruleofCourseCell["cell_struct"][i] == "period"
                periods::Vector{Vector{Int}} = ruleofCourseCell["syntax"]["period"](String(cell[i]))
            else
                throw(ArgumentError("invalid filed of ruleofCourseCell[\"cell_struct\"] valued $(ruleofCourseCell["cell_struct"][i])"))
            end
        end
    end
    deal_lesson_periods!(lessons, lesson, periods, term_begin_at)
    return lessons
end

"""
将山大学生学期课表(.xlsx)转换为符合钉钉日程api的Dict类型
有课单元格的格式和没课单元格的格式必须分别统一, 课程表范围内不要出现第三种格式
可选参数：
ruleofTable => 课程表的结构信息
ruleofCourseCell => 单元格的结构信息
default_struct => 日历api的课程表默认结构(对应的Dict)
"""
function xlsx2DingTalkDict(xlsx_file::String, configuration::Dict{String, Dict{String, Any}})
    rules_of_table          =   configuration["rules_of_table"]
    rules_of_course_cell    =   configuration["rules_of_course_cell"]
    api_struct              =   configuration["api_struct"]

    #read data from xlsx
    table = XLSX.readdata(xlsx_file, rules_of_table["work_sheet"], rules_of_table["index_ref"])

    #check range in one day
    if size(table)[1] > rules_of_table["course_struct"]["daily_course_amount"]
        throw(ArgumentError("size of table exceed daily_course_amount"))
    end

    #check if term is begin at Monday
    if dayofweek(rules_of_table["term_begin_at"][1]) != Monday
        throw(ArgumentError("term is not begin at Monday"))
    end

    #initialize lessons_set
    lessons_set = Set{Dict{String, Any}}()

    #traverse cell
    for day_in_week in 1 : size(table)[2]
        for sequence in 1 : size(table)[1]

            #set value of time offset
            t = Day(day_in_week + rules_of_table["week_begin_at"] - 2)
            se = rules_of_table["course_struct"]["course_periods"][sequence]

            #get cell
            cell = table[sequence, day_in_week]

            #ignore invalid cell
            if  !ismissing(cell)
                lessons = deal_course_cell(api_struct, cell, rules_of_table["term_begin_at"], rules_of_course_cell)

                #modify the time of lesson
                for lesson in lessons

                    #apply offset and convert to String type
                    lesson["start"]["dateTime"] = Dates.format(ZonedDateTime((lesson["start"]["dateTime"]  + t + se[1]), rules_of_table["term_begin_at"][2]), "yyyy-mm-ddTHH:MM:SSzzzz")
                    lesson["end"]["dateTime"] = Dates.format(ZonedDateTime((lesson["end"]["dateTime"]  + t + se[2]), rules_of_table["term_begin_at"][2]), "yyyy-mm-ddTHH:MM:SSzzzz")

                    #set arg for lesson which recurrence
                    if "recurrence" in keys(lesson)
                        lesson["recurrence"]["pattern"]["daysOfWeek"] = lowercase(dayname(day_in_week))
                    end

                    #push lessons into lessons_set
                    push!(lessons_set, deepcopy(lesson))
                end
            end
            println("good with day_in_week=$(day_in_week) and sequence=$(sequence))")
        end
    end

    return deepcopy(lessons_set)
end
end
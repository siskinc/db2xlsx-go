package main

import (
	"database/sql"
	"fmt"
	_ "github.com/Go-SQL-Driver/MySQL"
	"flag"
	"github.com/Luxurioust/excelize"
	"strconv"
	"strings"
	"bytes"
)

var str2xlsx map[string]string = map[string]string{
	"schoolId":  "学号",
	"name":      "姓名",
	"college":   "学院",
	"specialty": "专业",
	"grade":     "班级",
	"password":  "密码",
	"is_done":   "是否做题",
	"fraction":  "分数",
}

var colleges map[string]string = map[string]string{
	"1":"材料科学与工程学院",
	"2":"法学院" ,
	"3":"管理学院",
	"4":"化学工程学院",
	"5":"化学与环境工程学院",
	"6":"机械工程学院" ,
	"7":"计算机学院" ,
	"8":"教育与心理科学学院",
	"9":"经济学院" ,
	"10":"马克思主义学院",
	"11":"美术学院" ,
	"12":"人文学院" ,
	"13":"生物工程学院",
	"14":"数学与统计学院",
	"15":"体育学院" ,
	"16":"土木工程学院",
	"17":"外语学院" ,
	"18":"物理与电子工程学院",
	"19":"音乐学院" ,
	"20":"自动化与信息工程学院",
}

func main() {
	student := flag.Bool("student",false,"是否生成学生的表")
	teacher := flag.Bool("teacher",false,"是否生成教师的表")
	flag.Parse()
	options := map[string]bool{"student":*student,"teacher":*teacher}
	db, err := sql.Open("mysql", "yy:wyysdsa!@tcp(219.221.176.204:3306)/aqzsjs?charset=utf8")
	checkErr(err)
	defer db.Close()
	query, err := db.Query("SELECT * FROM aqzsjs.aq_member")
	checkErr(err)
	columns, results := getResult(query)
	toXlsx(columns,results,options)
}

func checkErr(errMsg error) {
	if errMsg != nil {
		panic(errMsg)
	}
}

func getResult(query *sql.Rows) (columns []string, result map[int]map[string]string) {
	column, _ := query.Columns()
	values := make([][]byte, len(column))
	scans := make([]interface{}, len(column))
	for i := range values {
		scans[i] = &values[i]
	}
	results := make(map[int]map[string]string)
	i := 0
	for query.Next() {
		if err := query.Scan(scans...); err != nil {
			fmt.Println(err)
			return
		}
		row := make(map[string]string)
		for k, v := range values {
			key := column[k]
			v = bytes.TrimRight(v,"\x00")
			row[key] = string(v)
			//fmt.Println(key,row[key])
		}
		results[i] = row
		i++
	}
	return column, results
}

func toXlsx(column []string, results map[int]map[string]string, options map[string]bool)  {
	student,teacher := filter(results)
	xlsx := excelize.NewFile()
	xlsx.DeleteSheet("Sheet1")
	if options["student"] == true {
		generateXlsxFileStudent(student,xlsx)
	}
	if options["teacher"] == true {
		generateXlsxFileTeacher(teacher,xlsx)
	}
	err := xlsx.SaveAs("./aqzsjs.xlsx")
	checkErr(err)
}

func generateXlsxFileTeacher(people map[int]map[string]string,xlsx *excelize.File) {
	index := xlsx.NewSheet("TeacherSheet")
	xlsx.SetCellValue("TeacherSheet", "A1", "工号")
	xlsx.SetCellValue("TeacherSheet", "B1", "姓名")
	xlsx.SetCellValue("TeacherSheet", "C1", "学院")
	xlsx.SetCellValue("TeacherSheet", "D1", "分数")
	xlsx.SetCellValue("TeacherSheet", "E1", "是否做题")
	i := 2
	for _,v := range people {
		xlsx.SetCellValue("TeacherSheet", "A"+strconv.Itoa(i), v["password"])
		xlsx.SetCellValue("TeacherSheet", "B"+strconv.Itoa(i), v["name"])
		_,err := strconv.ParseInt(v["college"],10,32)
		if err == nil {
			xlsx.SetCellValue("TeacherSheet", "C"+strconv.Itoa(i), colleges[v["college"]])
		} else {
			xlsx.SetCellValue("TeacherSheet", "C"+strconv.Itoa(i), strings.TrimSpace(v["college"]))
		}
		xlsx.SetCellValue("TeacherSheet", "D"+strconv.Itoa(i), v["fraction"])
		is_done,_ := strconv.ParseInt(v["is_done"],10,32)
		var is_done_str string
		//fmt.Println("is_done:",is_done)
		if is_done == 1 {
			is_done_str = "是"
		} else {
			is_done_str = "否"
		}
		xlsx.SetCellValue("TeacherSheet", "E"+strconv.Itoa(i), is_done_str)
		i++
 	}
	// Set active sheet of the workbook.
	xlsx.SetActiveSheet(index)
}

func generateXlsxFileStudent(people map[int]map[string]string,xlsx *excelize.File) {
	index := xlsx.NewSheet("StudentSheet")
	xlsx.SetCellValue("StudentSheet", "A1", "学号")
	xlsx.SetCellValue("StudentSheet", "B1", "姓名")
	xlsx.SetCellValue("StudentSheet", "C1", "学院")
	xlsx.SetCellValue("StudentSheet", "D1", "专业名称")
	xlsx.SetCellValue("StudentSheet", "E1", "行政班")
	xlsx.SetCellValue("StudentSheet", "F1", "身份证号")
	xlsx.SetCellValue("StudentSheet", "G1", "分数")
	xlsx.SetCellValue("StudentSheet", "H1", "是否做题")


	i := 2
	for _,v := range people {
		xlsx.SetCellValue("StudentSheet", "A"+strconv.Itoa(i), v["schoolId"])
		xlsx.SetCellValue("StudentSheet", "B"+strconv.Itoa(i), v["name"])
		_,err := strconv.ParseInt(v["college"],10,32)
		if err == nil {
			xlsx.SetCellValue("StudentSheet", "C"+strconv.Itoa(i), colleges[v["college"]])
		} else {
			xlsx.SetCellValue("StudentSheet", "C"+strconv.Itoa(i), v["college"])
		}
		xlsx.SetCellValue("StudentSheet", "D"+strconv.Itoa(i), v["specialty"])
		xlsx.SetCellValue("StudentSheet", "E"+strconv.Itoa(i), v["grade"])
		xlsx.SetCellValue("StudentSheet", "F"+strconv.Itoa(i), v["password"])
		xlsx.SetCellValue("StudentSheet", "G"+strconv.Itoa(i), v["fraction"])
		is_done,_ := strconv.ParseInt(v["id_done"],10,32)
		var is_done_str string
		if is_done == 1 {
			is_done_str = "是"
		} else {
			is_done_str = "否"
		}
		xlsx.SetCellValue("StudentSheet", "H"+strconv.Itoa(i), is_done_str)
		i++
	}
	// Set active sheet of the workbook.
	xlsx.SetActiveSheet(index)
}

func filter(results map[int]map[string]string) (students,teachers map[int]map[string]string) {
	student := make(map[int]map[string]string)
	teacher := make(map[int]map[string]string)
	i := 0
	j := 0
	for _,v := range results {
		if len(v["password"]) == 18 {
			student[i] = v
			i++
		} else {
			teacher[j] = v
			j++
		}
	}
	return student,teacher
}
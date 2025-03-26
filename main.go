package main

import (
	"errors"
	"flag"
	"fmt"
	"log"
	"os"
	"path"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

func usage() {
	fmt.Println("Usage: execl2csv [options] [execelName]")
	flag.PrintDefaults()
}

func sheetsArgsFormat(sheets string) (map[string]interface{}, error) {
	sss := make(map[string]interface{})
	if sheets == "" {
		return map[string]interface{}{}, nil
	} else {
		for _, sheet := range strings.Split(sheets, ";") {

			s := strings.Split(sheet, ":")
			if len(s) == 1 {
				sss[s[0]] = map[string]interface{}{}
			} else if len(s) == 2 {
				aa := make(map[string]interface{})
				ss := strings.Split(s[1], "->")
				if len(ss) == 1 || len(ss) == 2 {
					ss1 := strings.Split(ss[0], ",")
					if len(ss1) == 2 {
						x0, err := strconv.Atoi(ss1[0])
						if err != nil {
							return map[string]interface{}{}, errors.New(fmt.Sprintf("error: %s sheet 起始x坐标错误", s[0]))
						}

						y0, err := strconv.Atoi(ss1[1])
						if err != nil {
							return map[string]interface{}{}, errors.New(fmt.Sprintf("error: %s sheet 起始y坐标错误", s[0]))
						}
						aa["x0"] = x0
						aa["y0"] = y0
					} else {
						return map[string]interface{}{}, errors.New(fmt.Sprintf("error: %s sheet 起始坐标错误", s[0]))
					}
				}

				if len(ss) == 2 {
					ss1 := strings.Split(ss[1], ",")
					if len(ss1) == 2 {
						x1, err := strconv.Atoi(ss1[0])
						if err != nil {
							return map[string]interface{}{}, errors.New(fmt.Sprintf("error: %s sheet 结束x坐标错误", s[0]))
						}

						y1, err := strconv.Atoi(ss1[1])
						if err != nil {
							return map[string]interface{}{}, errors.New(fmt.Sprintf("error: %s sheet 结束y坐标错误", s[0]))
						}
						aa["x1"] = x1
						aa["y1"] = y1
						if aa["x0"].(int)*aa["y0"].(int)*aa["x1"].(int)*aa["y1"].(int) < 1 || aa["x0"].(int) > aa["x1"].(int) || aa["y0"].(int) > aa["y1"].(int) {
							return map[string]interface{}{}, errors.New(fmt.Sprintf("error: %s sheet 结束坐标大于起始坐标", s[0]))
						}
					} else {
						return map[string]interface{}{}, errors.New(fmt.Sprintf("error: %s sheet 结束坐标错误", s[0]))
					}

				}
				if len(ss) > 2 || len(ss) < 1 {
					return map[string]interface{}{}, errors.New(fmt.Sprintf("error: %s sheet 坐标参数错误", s[0]))
				}

				sss[s[0]] = aa
			}
		}
		return sss, nil
	}

}

func main() {
	var ft = flag.String("ft", ",", "字段分割符")
	var ec = flag.String("ec", "", "字段包围符")
	var lt = flag.String("lt", "\r\n", "行分隔符")
	var out = flag.String("o", "", "csv保存路径")
	var ig = flag.String("ig", "0", "跳过前几行")
	var sheets = flag.String("sheet", "", "指定sheet和范围(行，列)，如不指定范围则默认sheet内容全部转换，格式:sheet1:1,2->4,8;sheet2:1,1->8,8;sheet3:2,2->16,16")

	flag.Usage = usage
	flag.Parse()
	args := flag.Args()
	if len(args) == 0 {
		usage()
		return
	}
	log.SetFlags(log.Ldate | log.Ltime | log.Lshortfile)
	skipNum, err := strconv.Atoi(*ig)
	if err != nil {
		log.Printf("error: ig参数错误，%s", err.Error())
		return
	}

	var sss map[string]interface{}
	if *sheets != "" {
		sss, err = sheetsArgsFormat(*sheets)
		if err != nil {
			log.Println(err.Error())
			return
		}
	}

	input := args[0]
	f, err := excelize.OpenFile(input)
	defer f.Close()

	if err != nil {
		log.Println("打开", input, "失败：", err.Error())
		return
	}

	for _, name := range f.GetSheetMap() {
		rows, err := f.GetRows(name)
		if err != nil {
			log.Println("获取sheet数据失败", err.Error())
			return
		}
		cols, err := f.GetCols("Sheet1")
		if err != nil {
			log.Println("获取sheet数据失败", err.Error())
			return
		}
		table_length := len(rows)
		table_width := len(cols)

		var x0, y0, x1, y1 int

		if position, ok := sss[name]; !ok {
			if len(sss) > 0 {
				continue
			}
			if skipNum > 0 {
				x0 = skipNum
			} else {
				x0 = 0
			}
			y0 = 0
			x1 = table_length
			y1 = table_width
		} else if ok {
			pos := position.(map[string]interface{})
			if len(pos) == 2 {
				x0 = pos["x0"].(int) - 1
				y0 = pos["y0"].(int) - 1
				x1 = table_length
				y1 = table_width
			}

			if len(pos) == 4 {
				x0 = pos["x0"].(int) - 1
				y0 = pos["y0"].(int) - 1
				x1 = pos["x1"].(int)
				if x1 > table_length {
					x1 = table_length
				}
				y1 = pos["y1"].(int)
				if y1 > table_width {
					y1 = table_width
				}
			}
		}
		s1 := ""
		for _, row := range rows[x0:x1] {
			s2 := ""
			for i := y0; i < y1; i++ {

				cell := ""
				if i < len(row) {
					cell = row[i]
				}
				if s2 == "" {
					s2 = fmt.Sprint(*ec, cell, *ec)
				} else {
					s2 += fmt.Sprint(*ft, *ec, cell, *ec)
				}
			}
			s1 += fmt.Sprint(s2, *lt)
		}

		csvpath := path.Join(*out, name+".csv")
		outfile, err := os.OpenFile(csvpath, os.O_WRONLY|os.O_CREATE, 0644)
		defer outfile.Close()
		if err != nil {
			log.Println("打开输出文件失败", err.Error())
			return
		}

		outfile.WriteString(s1)
	}

}

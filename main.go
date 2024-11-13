package main

import (
	"flag"
	"fmt"
	"log"
	"os"

	"github.com/xuri/excelize/v2"
)

func usage() {
	fmt.Println("Usage: execl2csv [options] [execelName]")
	flag.PrintDefaults()
}

func main() {
	var ft = flag.String("ft", ",", "字段分割符")
	var ec = flag.String("ec", "\"", "字段包围符")
	var lt = flag.String("lt", "\r\n", "行分隔符")
	var out = flag.String("o", "", "csv保存路径")
	var nh = flag.String("nh", "0", "是否忽略表头,默认忽略,1表示忽略，0表示不忽略")
	flag.Usage = usage
	flag.Parse()
	args := flag.Args()
	if len(args) == 0 {
		usage()
		return
	}
	input := args[0]
	f, err := excelize.OpenFile(input)
	defer f.Close()

	if err != nil {
		log.Println("打开", input, "失败：", err.Error())
		return
	}

	for _, name := range f.GetSheetMap() {
		s1 := ""
		rows, err := f.GetRows(name)
		if err != nil {
			log.Println("获取sheet数据失败", err.Error())
			return
		}
		for row_id, row := range rows {
			if *nh == "1" && row_id == 0 {
				continue
			}
			s2 := ""
			for _, cell := range row {
				if s2 == "" {
					s2 = fmt.Sprint(*ec, cell, *ec)
				} else {
					s2 += fmt.Sprint(*ft, *ec, cell, *ec)
				}
			}
			s1 += fmt.Sprint(s2, *lt)
		}
		outfile, err := os.OpenFile(*out+name+".csv", os.O_WRONLY|os.O_CREATE, 0644)
		defer outfile.Close()
		if err != nil {
			log.Println("打开输出文件失败", err.Error())
			return
		}

		outfile.WriteString(s1)
	}

}

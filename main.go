package main

import (
	"flag"
	"fmt"
	"log"
	"os"

	"github.com/xuri/excelize/v2"
)

func usage() {
	fmt.Println(`Usage: execl2csv [options] [execelName]
  -ec string
        字段包围符 (default ")
  -ft string
        字段分割符 (default ,)
  -lt string
        行分隔符 (default \r\n)
  -o string
        csv保存路径`)
}

func main() {
	var ft = flag.String("ft", ",", "字段分割符")
	var ec = flag.String("ec", "\"", "字段包围符")
	var lt = flag.String("lt", "\r\n", "行分隔符")
	var out = flag.String("o", "", "csv保存路径")
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
		cols, err := f.GetCols(name)
		if err != nil {
			log.Println("获取sheet数据失败", err.Error())
			return
		}
		for _, col := range cols {
			s2 := ""
			for _, rowCell := range col {
				if s2 == "" {
					s2 = fmt.Sprint(*ec, rowCell, *ec)
				} else {
					s2 += fmt.Sprint(*ft, *ec, rowCell, *ec)
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

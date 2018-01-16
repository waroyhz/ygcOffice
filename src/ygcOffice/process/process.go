package process

import (
	"github.com/larspensjo/config"
	"ygcOffice/define"
	"strings"
	"ygcOffice/excel"
	"fmt"
	"time"
	"strconv"
	"reflect"
	"github.com/Luxurioust/excelize"
	"errors"
)

type YgcOfficeProcess struct {
	target *excel.ExcelObject
	//readData  []interface{}
	data map[string]interface{}
}

const (
	key_data   = "$data"
	key_xstart = "$xStartText"
	key_ystart = "$yStartText"
	key_xend   = "$xEndText"
	key_yend   = "$yEndText"
)

type ColumnMap struct {
	SrcColumn string
	DstColumn string
	Value     string
}

func (this *YgcOfficeProcess) setReadData(data []interface{}) {
	this.data[key_data] = data
}

func (this *YgcOfficeProcess) getReadData() []interface{} {
	if val, ok := this.data[key_data]; ok {
		return val.([]interface{})
	} else {
		return []interface{}{}
	}
}

func NewProcess(cfg *config.Config, section string, srcxlsx, dstxlsx *excelize.File) {
	var data = map[string]interface{}{}
	data[define.KEY_VALUE_src] = &excel.ExcelObject{File: srcxlsx}
	data[define.KEY_VALUE_dst] = &excel.ExcelObject{File: dstxlsx}
	data[define.KEY_VALUE_compny] = excel.GetCompnyNameFromPath(srcxlsx.Path)
	var ygcOfficeProcess = YgcOfficeProcess{}
	ygcOfficeProcess.data = data
	ygcOfficeProcess.ProcessNext(cfg, section)
}

func (this *YgcOfficeProcess) ProcessNext(cfg *config.Config, section string, ) {
	if nextSection, err := cfg.String(section, define.KEY_OPTION_nextSection); err == nil {
		sections := strings.Split(nextSection, ",")
		for _, v := range sections {
			this.ProcessCommand(cfg, strings.Trim(v, " "))
		}
	}
}
func (this *YgcOfficeProcess) ProcessCommand(cfg *config.Config, section string) {
	fmt.Printf("处理节点：%s\n", section, )
	if target, err := cfg.String(section, define.KEY_OPTION_target); err == nil {
		this.ProcessTarget(target)
	}
	if sheet, err := cfg.String(section, define.KEY_OPTION_sheet); err == nil {
		this.ProcessSheet(sheet)
	}
	if xstart, err := cfg.String(section, define.KEY_OPTION_xStartText); err == nil {
		xstart = findByMap(&this.data, xstart)
		this.ProcessXStartText(xstart)
	}
	if ystart, err := cfg.String(section, define.KEY_OPTION_yStartText); err == nil {
		ystart = findByMap(&this.data, ystart)
		this.ProcessYStartText(ystart)
	}

	if xstart, err := cfg.String(section, define.KEY_OPTION_xFindText); err == nil {
		xstart = findByMap(&this.data, xstart)
		this.ProcessXFindText(xstart)
	}
	if ystart, err := cfg.String(section, define.KEY_OPTION_yFindText); err == nil {
		ystart = findByMap(&this.data, ystart)
		this.ProcessYFindText(ystart)
	}
	if o, err := cfg.String(section, define.KEY_OPTION_operation); err == nil {
		this.target.Operation = o
	}
	if xend, err := cfg.String(section, define.KEY_OPTION_xEndText); err == nil {
		xend = findByMap(&this.data, xend)
		this.ProcessXEndText(xend)
	}
	if yend, err := cfg.String(section, define.KEY_OPTION_yEndText); err == nil {
		yend = findByMap(&this.data, yend)
		this.ProcessYEndText(yend)
	}
	if xadd, err := cfg.Int(section, define.KEY_OPTION_xAdd); err == nil {
		this.ProcessXAdd(xadd)
	}
	if yadd, err := cfg.Int(section, define.KEY_OPTION_yAdd); err == nil {
		this.ProcessYAdd(yadd)
	}

	if cfg.HasOption(section, define.KEY_OPTION_process) {
		this.ProcessOption(cfg, section)
	}

	this.ProcessNext(cfg, section)
}
func findByMap(data *map[string]interface{}, key string) string {
	if len(key) > 0 && key[0:1] == "$" {
		if v, ok := (*data)[key]; ok {
			if key == key_data {
				return key
			}
			return v.(string)
		}
	}
	return key
}

func (this *YgcOfficeProcess) ProcessYAdd(yadd int) {
	this.target.Y += yadd
	fmt.Printf("设置当前坐标：%s\n", excel.GetCellName(this.target.X, this.target.Y))
}
func (this *YgcOfficeProcess) ProcessXAdd(xadd int) {
	this.target.X += xadd
	fmt.Printf("设置当前坐标：%s\n", excel.GetCellName(this.target.X, this.target.Y))
}

func (this *YgcOfficeProcess) ProcessOption(cfg *config.Config, section string) {
	if process, err := cfg.String(section, define.KEY_OPTION_process); err == nil {
		if process == define.KEY_VALUE_readArray {
			this.processReadArray(cfg, section)
		} else if process == define.KEY_VALUE_writeArray {
			this.processWriteArray(cfg, section)
		} else if process == define.KEY_VALUE_readItemProcess {
			this.processItemProcess(cfg, section)
		} else if process == define.KEY_VALUE_readItem {
			this.processReadItem(cfg, section)
		} else {
			panic("process 读取操作报错：" + process)
		}
	} else {
		panic("读取操作报错：" + section)
	}
}

func (this *YgcOfficeProcess) processItemProcess(cfg *config.Config, section string) {
	var readEndCondition string
	if e, err := cfg.String(section, define.KEY_OPTION_readEndCondition); err == nil {
		readEndCondition = e
	} else {
		readEndCondition = "==========================================================================================" //这个字符不会被匹配
	}
	var readRange int
	if e, err := cfg.Int(section, define.KEY_OPTION_readRange); err == nil {
		readRange = e
	}

	var processSection []string
	if ps, err := cfg.String(section, define.KEY_OPTION_processSection); err == nil {
		processSection = strings.Split(ps, ",")
	} else {
		panic("processItemProcess 必须指定 processSection 处理节点")
	}

	columnMap := []*ColumnMap{}
	if cmap, err := cfg.String(section, define.KEY_OPTION_columnMap); err == nil {
		smap := strings.Split(cmap, ",")
		for _, s := range smap {
			t := strings.Split(s, "=")
			columnMap = append(columnMap, &ColumnMap{SrcColumn: t[0], DstColumn: t[1]})
		}
	}

	psrc := this.data[define.KEY_VALUE_src].(*excel.ExcelObject)
	src := *psrc
	pdst := this.data[define.KEY_VALUE_dst].(*excel.ExcelObject)
	//dst := *pdst

	this.setReadData([]interface{}{})

	if this.target.Operation == "" {
		fmt.Printf("节点 %s 未指定 Operation ，设置为默认 right", section)
	}
	if this.target.Operation == define.KEY_VALUE_right {
		fmt.Printf("向右读数据，当前坐标位置：%s，处理范围： %s-%s\n", excel.GetCellName(this.target.X, this.target.Y), excel.GetCellName(this.target.X, this.target.Y), excel.GetCellName(this.target.Xe, this.target.Ye))
		for i := this.target.X; i < this.target.Xe; i++ {
			val := this.target.File.GetCellValue(this.target.CurrentSheet, excel.GetCellName(i, this.target.Y))
			if val == readEndCondition {
				break
			}
			if readRange > 0 && this.target.Xe-i > readRange {
				break
			}
			src.X = i
			for _, v := range processSection {
				createNewChildProcessObject(cfg, v, src, pdst, this.target, this.getReadData())
			}
			pdst.X++
		}
	} else if this.target.Operation == define.KEY_VALUE_down {
		fmt.Printf("向下读数据，当前坐标位置：%s，处理范围： %s-%s\n", excel.GetCellName(this.target.X, this.target.Y), excel.GetCellName(this.target.X, this.target.Y), excel.GetCellName(this.target.Xe, this.target.Ye))
		for i := this.target.Y; i < this.target.Ye; i++ {
			val := this.target.File.GetCellValue(this.target.CurrentSheet, excel.GetCellName(this.target.X, i))
			if val == readEndCondition {
				break
			}
			if readRange > 0 && this.target.Ye-i > readRange {
				break
			}
			src.Y = i
			if len(columnMap) > 0 {
				//生成源key
				var key string
				this.ProcessTarget(define.KEY_VALUE_src)
				for _, cmap := range columnMap {
					this.ProcessXFindText(cmap.SrcColumn)
					cmap.Value = this.target.File.GetCellValue(this.target.CurrentSheet, excel.GetCellName(this.target.X, i))
					if len(key) > 0 {
						key += ","
					}
					key += cmap.Value
				}
				this.ProcessTarget(define.KEY_VALUE_dst)
				find := false
				insert := false
				//读取目标key
				for di := this.target.Ys; di < this.target.Ye; di++ {
					dstKey := ""
					for _, cmap := range columnMap {
						this.ProcessXFindText(cmap.DstColumn)
						val := this.target.File.GetCellValue(this.target.CurrentSheet, excel.GetCellName(this.target.X, di))
						if len(dstKey) > 0 {
							dstKey += ","
						}
						dstKey += val
					}
					if dstKey == key {
						find = true
					}
					if dstKey == "" {
						this.target.Y = di
						insert = true
						break
					}
				}
				if find {
					fmt.Printf("map 找到重复数据 " + key)
				} else if insert {
					for _, cmap := range columnMap {
						this.ProcessXFindText(cmap.DstColumn)
						val := this.target.File.GetCellValue(this.target.CurrentSheet, excel.GetCellName(this.target.X, this.target.Y))
						this.writeNewValue(val, cmap.Value, "", excel.GetCellName(this.target.X, this.target.Y))
					}
				} else {
					panic("没有根据map找到可插入的位置 " + key)
				}
				this.ProcessTarget(define.KEY_VALUE_src)
			}
			for _, v := range processSection {
				createNewChildProcessObject(cfg, v, src, pdst, this.target, this.getReadData())
			}
			pdst.Y++
		}
	}
}

func (this *YgcOfficeProcess) processWriteArray(cfg *config.Config, section string) {
	var format string
	if e, err := cfg.String(section, define.KEY_OPTION_format); err == nil {
		format = e
	}

	var override string
	if e, err := cfg.String(section, define.KEY_OPTION_hasData); err == nil {
		override = e
	}

	fmt.Printf("准备写数据，当前坐标位置：%s 处理范围： %s-%s\n", excel.GetCellName(this.target.X, this.target.Y), excel.GetCellName(this.target.X, this.target.Y), excel.GetCellName(this.target.Xe, this.target.Ye))

	if len(this.getReadData()) == 0 {
		this.setReadData(append(this.getReadData(), ""))
	}
	for i, v := range this.getReadData() {
		if this.target.Operation == "" || this.target.Operation == define.KEY_VALUE_sum {
			val := this.target.File.GetCellValue(this.target.CurrentSheet, excel.GetCellName(this.target.X, this.target.Y))
			if override != define.KEY_VALUE_sum && val != "" {
				//是否覆盖数据？
				val = ""
			}
			this.writeNewValue(val, v, format, excel.GetCellName(this.target.X, this.target.Y))
		} else if this.target.Operation == define.KEY_VALUE_right {
			//for i := this.target.X; i < this.target.Xe; i++ {
			val := this.target.File.GetCellValue(this.target.CurrentSheet, excel.GetCellName(this.target.X+i, this.target.Y))
			if val != "" {
				val = ""
			}
			this.writeNewValue(val, v, format, excel.GetCellName(this.target.X+i, this.target.Y))
			//}
		} else if this.target.Operation == define.KEY_VALUE_down {
			//for i := this.target.Y; i < this.target.Ye; i++ {
			val := this.target.File.GetCellValue(this.target.CurrentSheet, excel.GetCellName(this.target.X, this.target.Y+i))
			if override != define.KEY_VALUE_sum && val != "" {
				val = ""
			}
			this.writeNewValue(val, v, format, excel.GetCellName(this.target.X, this.target.Y+i))
			//}
		}
	}
	this.setReadData([]interface{}{})
}
func (this *YgcOfficeProcess) writeNewValue(oldval string, newval interface{}, format string, cellName string) {
	var wValue interface{}
	var style = 0
	if format == "" {
		fmt.Printf("未指定 format , 默认使用 string ")
		format = define.KEY_VALUE_string
	}
	switch format {
	case define.KEY_VALUE_string:
		nv := fmt.Sprintf("%v", newval)
		wValue = oldval + nv
	case define.KEY_VALUE_date:
		value := fmt.Sprintf("%v", newval)
		dateValue := "2006-01-02T00:00:00+08:00"
		//8位日期
		dateValue = strings.Replace(dateValue, "2006", "20"+value[6:], -1)
		dateValue = strings.Replace(dateValue, "-01-", fmt.Sprintf("-%s-", value[0:2]), -1)
		dateValue = strings.Replace(dateValue, "02T", fmt.Sprintf("%sT", value[3:5]), -1)
		t, err := time.Parse(time.RFC3339, dateValue)
		if dateValue != "" && err != nil {
			panic(err)
		}
		wValue = t
		//style, _ = this.target.File.NewStyle(`{"custom_number_format": "yyyy/mm/dd"}`)
		//style, err := xlsx.NewStyle(`{"border":[{"type":"left","color":"0000FF","style":3},{"type":"top","color":"00FF00","style":4},{"type":"bottom","color":"FFFF00","style":5},{"type":"right","color":"FF0000","style":6},{"type":"diagonalDown","color":"A020F0","style":7},{"type":"diagonalUp","color":"A020F0","style":8}]}`)
		style, _ = this.target.File.NewStyle(`{"border":[{"type":"left","color":"000000","style":1},{"type":"top","color":"000000","style":1},{"type":"bottom","color":"000000","style":1},{"type":"right","color":"000000","style":1}],"custom_number_format": "yyyy/mm/dd"}`)
		//style,_=this.target.File.NewStyle(`{"number_format": 22}`)

	case define.KEY_VALUE_float:
		if oldval == "" {
			oldval = "0"
		}
		if oldvalue, err := strconv.ParseFloat(oldval, 64); err == nil {
			switch reflect.ValueOf(newval).Kind() {
			case reflect.Float64, reflect.Float32:
				wValue = oldvalue + reflect.ValueOf(newval).Float()
			case reflect.Int, reflect.Int32, reflect.Int64, reflect.Int8:
				wValue = oldvalue + float64(reflect.ValueOf(newval).Int())
			case reflect.String:
				if len(newval.(string)) == 0 {
					newval = "0"
				}
				if n, e := strconv.ParseFloat(newval.(string), 64); e == nil {
					wValue = oldvalue + n
				} else {
					panic(e)
				}
			default:
				panic("format 未处理的类型 ")
			}
			style, _ = this.target.File.NewStyle(`{"border":[{"type":"left","color":"000000","style":1},{"type":"top","color":"000000","style":1},{"type":"bottom","color":"000000","style":1},{"type":"right","color":"000000","style":1}],"number_format": 43}`)

		} else {
			panic(err)
		}

		//case define.KEY_VALUE_int:
	default:
		panic("format 未处理的类型 " + format)
	}
	this.target.File.SetCellValue(this.target.CurrentSheet, cellName, wValue)
	if style > 0 {
		this.target.File.SetCellStyle(this.target.CurrentSheet, cellName, cellName, style)
	}
	fmt.Printf("写入数据：%s  %v\n", cellName, wValue)
}

//func processSum() {
//	var val float64 =0
//	for _,v:=range readData{
//		switch reflect.ValueOf(v).Type().Kind(){
//		case reflect.String:
//			if t,e:= strconv.ParseFloat(v.(string),64);e==nil{
//				val+=t
//			}
//		case reflect.Float64, reflect.Float32:
//			val+=v.(float64)
//		default:
//			panic("未处理的类型 "+ reflect.ValueOf(v).Type().Kind().String())
//		}
//	}
//	readData = []interface{}{}
//	readData = append(readData, val)
//}

func (this *YgcOfficeProcess) processReadArray(cfg *config.Config, section string) {
	var readEndCondition string
	if e, err := cfg.String(section, define.KEY_OPTION_readEndCondition); err == nil {
		readEndCondition = e
	} else {
		readEndCondition = "==========================================================================================" //这个字符不会被匹配
	}
	var readRange int
	if e, err := cfg.Int(section, define.KEY_OPTION_readRange); err == nil {
		readRange = e
	}

	this.setReadData([]interface{}{})
	fmt.Printf("准备读数据，当前坐标位置：%s，处理范围： %s-%s\n", excel.GetCellName(this.target.X, this.target.Y), excel.GetCellName(this.target.Xe, this.target.Ye))
	if this.target.Operation == "" {
		this.setReadData(append(this.getReadData(), this.target.File.GetCellValue(this.target.CurrentSheet, excel.GetCellName(this.target.X, this.target.Y))))
		fmt.Printf("读到数据：%v\n", this.getReadData())

	} else if this.target.Operation == define.KEY_VALUE_right {
		for i := this.target.X; i < this.target.Xe; i++ {
			val := this.target.File.GetCellValue(this.target.CurrentSheet, excel.GetCellName(i, this.target.Y))
			if val == readEndCondition {
				break
			}
			if readRange > 0 && this.target.Xe-i > readRange {
				break
			}
			this.setReadData(append(this.getReadData(), val))
		}
	} else if this.target.Operation == define.KEY_VALUE_down {
		for i := this.target.Y; i < this.target.Ye; i++ {
			val := this.target.File.GetCellValue(this.target.CurrentSheet, excel.GetCellName(this.target.X, i))
			if val == readEndCondition {
				break
			}
			if readRange > 0 && this.target.Ye-i > readRange {
				break
			}
			this.setReadData(append(this.getReadData(), val))
		}
	}
	fmt.Printf("读到数据：%v\n", this.getReadData())
}

func (this *YgcOfficeProcess) ProcessXEndText(s string) {
	isFindData := false
	if s == key_data {
		s = this.getReadData()[0].(string)
		isFindData = true
	}
	if s == key_xstart {
		this.target.Xe = this.target.X
	} else if len(s) > 0 && s[0:1] == "$" {
		panic(errors.New(fmt.Sprintf("参数错误 %s", s)))
	} else if result, fx := excel.FindConlumnCell(this.target.File, this.target.CurrentSheet, this.target.X, this.target.X+1000, this.target.Y, s); result {
		this.target.Xe = fx
	} else {
		if isFindData {
			if result, fx := excel.FindConlumnCell(this.target.File, this.target.CurrentSheet, this.target.X, this.target.X+1000, this.target.Y, ""); result {
				this.target.Xe = fx
			} else {
				panic("找不到字符串 " + s)
			}
		} else {
			panic("找不到字符串 " + s)
		}
	}
	fmt.Printf("设置处理范围： %s-%s\n", excel.GetCellName(this.target.Xs, this.target.Ys), excel.GetCellName(this.target.Xe, this.target.Ye))
}
func (this *YgcOfficeProcess) ProcessYEndText(s string) {
	isFindData := false
	if s == key_data {
		s = this.getReadData()[0].(string)
		isFindData = true
	}
	if s == key_ystart {
		this.target.Ye = this.target.Y
	} else if len(s) > 0 && s[0:1] == "$" {
		panic(errors.New(fmt.Sprintf("参数错误 %s", s)))
	} else if result, fy := excel.FindRowCell(this.target.File, this.target.CurrentSheet, this.target.X, this.target.Y, this.target.Y+5000, s); result {
		this.target.Ye = fy
	} else {
		if isFindData {
			if result, fy := excel.FindRowCell(this.target.File, this.target.CurrentSheet, this.target.X, this.target.Y, this.target.Y+5000, ""); result {
				this.target.Ye = fy
			} else {
				panic("找不到字符串 " + s)
			}
		} else {
			panic("找不到字符串 " + s)
		}
	}
	fmt.Printf("设置处理范围： %s-%s\n", excel.GetCellName(this.target.Xs, this.target.Ys), excel.GetCellName(this.target.Xe, this.target.Ye))
}

func (this *YgcOfficeProcess) ProcessYFindText(s string) {
	if s == key_data {
		s = this.getReadData()[0].(string)
	}
	if s == key_yend {
		this.target.Y = this.target.Ye
	} else if len(s) > 0 && s[0:1] == "$" {
		panic(errors.New(fmt.Sprintf("参数错误 %s", s)))
	} else if result, _, fy := excel.FindStartTextCell(this.target.File, this.target.CurrentSheet, 0, this.target.Ys, s); result {
		this.target.Y = fy
	} else {
		panic("找不到字符串 " + s)
	}
	//fmt.Printf("设置当前坐标：%s\n", excel.GetCellName(this.target.X, this.target.Y))
}
func (this *YgcOfficeProcess) ProcessXFindText(s string) {
	if s == key_data {
		s = this.getReadData()[0].(string)
	}
	if s == key_xend {
		this.target.X = this.target.Xe
	} else if len(s) > 0 && s[0:1] == "$" {
		panic(errors.New(fmt.Sprintf("参数错误 %s", s)))
	} else if result, fx, _ := excel.FindStartTextCell(this.target.File, this.target.CurrentSheet, 0, this.target.Ys, s); result {
		this.target.X = fx
	} else {
		panic("找不到字符串 " + s)
	}
	//fmt.Printf("设置当前坐标：%s\n", excel.GetCellName(this.target.X, this.target.Y))
}

func (this *YgcOfficeProcess) ProcessYStartText(s string) {
	if s == key_data {
		s = this.getReadData()[0].(string)
	}
	if s == key_yend {
		this.target.Ys = this.target.Ye
	} else if len(s) > 0 && s[0:1] == "$" {
		panic(errors.New(fmt.Sprintf("参数错误 %s", s)))
	} else if result, _, fy := excel.FindStartTextCell(this.target.File, this.target.CurrentSheet, this.target.Xs, this.target.Ys, s); result {
		this.target.Ys = fy
	} else {
		panic("找不到字符串 " + s)
	}

	if this.target.Y < this.target.Ys {
		this.target.Y = this.target.Ys
	}
	fmt.Printf("设置当前坐标：%s\n", excel.GetCellName(this.target.X, this.target.Y))
}
func (this *YgcOfficeProcess) ProcessXStartText(s string) {
	if s == key_data {
		s = this.getReadData()[0].(string)
	}
	if s == key_xend {
		this.target.Xs = this.target.Xe
	} else if len(s) > 0 && s[0:1] == "$" {
		panic(errors.New(fmt.Sprintf("参数错误 %s", s)))
	} else if result, fx, _ := excel.FindStartTextCell(this.target.File, this.target.CurrentSheet, this.target.Xs, this.target.Ys, s); result {
		this.target.Xs = fx
	} else {
		panic("找不到字符串 " + s)
	}
	if this.target.X < this.target.Xs {
		this.target.X = this.target.Xs
	}
	fmt.Printf("设置当前坐标：%s\n", excel.GetCellName(this.target.X, this.target.Y))
}

func (this *YgcOfficeProcess) ProcessSheet(sheet string) {
	this.target.CurrentSheet = sheet
}
func (this *YgcOfficeProcess) ProcessTarget(value string) {
	this.target = this.data[value].(*excel.ExcelObject)
}
func (this *YgcOfficeProcess) processReadItem(i *config.Config, s string) {
	pos := excel.GetCellName(this.target.X, this.target.Y)
	val := this.target.File.GetCellValue(this.target.CurrentSheet, pos)
	this.setReadData(append(this.getReadData(), val))
	fmt.Printf("在 %s 读到数据：%v\n", pos, this.getReadData())
}

//func (this *YgcOfficeProcess) processWriteItem(cfg *config.Config, section string) {
//	var format string
//	if e, err := cfg.String(section, define.KEY_OPTION_format); err == nil {
//		format = e
//	}
//
//	val := this.target.File.GetCellValue(this.target.CurrentSheet, excel.GetCellName(this.target.X, this.target.Y))
//	this.writeNewValue(val, v, format, excel.GetCellName(this.target.X, this.target.Y))
//}

func createNewChildProcessObject(cfg *config.Config, section string, src excel.ExcelObject, dst *excel.ExcelObject, target *excel.ExcelObject, readData []interface{}) {
	var data = map[string]interface{}{}
	data[define.KEY_VALUE_src] = &src
	data[define.KEY_VALUE_dst] = dst
	data[define.KEY_VALUE_compny] = excel.GetCompnyNameFromPath(src.File.Path)
	var ygcOfficeProcess = YgcOfficeProcess{}
	ygcOfficeProcess.data = data
	ygcOfficeProcess.setReadData(readData)
	ygcOfficeProcess.target = target
	ygcOfficeProcess.ProcessCommand(cfg, section)
}
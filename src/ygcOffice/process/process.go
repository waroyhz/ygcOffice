package process

import (
	"github.com/larspensjo/config"
	"ygcOffice/define"
	"strings"
	"ygcOffice/excel"
	"fmt"
	"strconv"
	"reflect"
	"github.com/360EntSecGroup-Skylar/excelize"
	"errors"
	"sort"
	"ygcOffice/catch"
	"os"
)

type YgcOfficeProcess struct {
	target *excel.ExcelObject
	data   map[string]interface{}
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

type TableMap struct {
	name             string
	srcColumn        string
	dstColumn        string
	format           string
	key              bool
	readEndCondition string

	srcX  int
	dstX  int
	value string
}
type TableMapValue struct {
	table *TableMap
	dstY  int
	srcY  int
	value interface{}
}

type TableMapLineData struct {
	value    interface{}
	lineData map[string]TableMapValue
}

var tableMap []*TableMap

func (this *YgcOfficeProcess) setReadData(data []TableMapLineData) {
	this.data[key_data] = data
}

func (this *YgcOfficeProcess) getReadData() []TableMapLineData {
	if val, ok := this.data[key_data]; ok {
		return val.([]TableMapLineData)
	} else {
		return []TableMapLineData{}
	}
}

func NewProcess(cfg *config.Config, section string, srcxlsx, dstxlsx *excelize.File, parentSection string) {
	var data = map[string]interface{}{}
	data[define.KEY_VALUE_src] = &excel.ExcelObject{File: srcxlsx}
	data[define.KEY_VALUE_dst] = &excel.ExcelObject{File: dstxlsx}
	data[define.KEY_VALUE_compny] = excel.GetCompnyNameFromPath(srcxlsx.Path)
	var ygcOfficeProcess = YgcOfficeProcess{}
	defer catch.Catch(func(param ... interface{}) {
		ygcOfficeProcess:= param[0].(*YgcOfficeProcess)
		fmt.Printf(`========================================================
当前正在操作文件：%s
程序发生意料之外的异常，无法继续执行，请仔细查看错误信息后按任意键退出……
`,ygcOfficeProcess.target.File.Path)
		var onkey string
		fmt.Scanln(onkey)
		os.Exit(1)
	},&ygcOfficeProcess)
	ygcOfficeProcess.data = data
	ygcOfficeProcess.ProcessNext(cfg, section, parentSection)
}

func (this *YgcOfficeProcess) ProcessNext(cfg *config.Config, section string, parentSection string) {
	if nextSection, err := cfg.String(section, define.KEY_OPTION_nextSection); err == nil {
		sections := strings.Split(nextSection, ",")
		for _, v := range sections {
			this.ProcessCommand(cfg, strings.Trim(v, " "), parentSection+"→"+section)
		}
	}
}
func (this *YgcOfficeProcess) ProcessCommand(cfg *config.Config, section string, parentSection string) {
	fmt.Printf("处理节点：%s→%s\n", parentSection, section)
	if _, err := cfg.Options(section); err != nil {
		panic(fmt.Sprintf("配置错误，节点 %s 不存在！", section))
	}
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
	if cfg.HasOption(section, define.KEY_OPTION_tableMap) {
		this.ProcessTableMap(cfg, section)
	}

	if cfg.HasOption(section, define.KEY_OPTION_process) {
		this.ProcessOption(cfg, section, parentSection)
	}

	this.ProcessNext(cfg, section, parentSection)
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
	fmt.Printf("设置当前Y坐标：%s\n", excel.GetCellName(this.target.X, this.target.Y))
}
func (this *YgcOfficeProcess) ProcessXAdd(xadd int) {
	this.target.X += xadd
	fmt.Printf("设置当前X坐标：%s\n", excel.GetCellName(this.target.X, this.target.Y))
}

func (this *YgcOfficeProcess) ProcessOption(cfg *config.Config, section string, parentSection string) {
	if process, err := cfg.String(section, define.KEY_OPTION_process); err == nil {
		if process == define.KEY_VALUE_readArray {
			this.processReadArray(cfg, section)
		} else if process == define.KEY_VALUE_writeArray {
			this.processWriteArray(cfg, section)
		} else if process == define.KEY_VALUE_readItemProcess {
			this.processItemProcess(cfg, section, parentSection)
		} else if process == define.KEY_VALUE_readItem {
			this.processReadItem(cfg, section)
		} else if process == define.KEY_VALUE_sum {
			this.processSum(cfg, section)
		} else if process == define.KEY_VALUE_sort {
			this.processSort(cfg, section)
		} else if process == define.KEY_VALUE_limt {
			this.processLimt(cfg, section)
		} else if process == define.KEY_VALUE_reset {
			this.processReset(cfg, section)
		} else if process == define.KEY_VALUE_filter {
			this.processFilter(cfg, section)
		} else {
			panic("process 读取操作报错：" + process)
		}
	} else {
		panic("读取操作报错：" + section)
	}
}

func (this *YgcOfficeProcess) processItemProcess(cfg *config.Config, section string, parentSection string) {
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

	this.setReadData([]TableMapLineData{})

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
				createNewChildProcessObject(cfg, v, src, pdst, this.target, this.getReadData(), parentSection)
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
				createNewChildProcessObject(cfg, v, src, pdst, this.target, this.getReadData(), parentSection)
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

	var readRange int
	if e, err := cfg.Int(section, define.KEY_OPTION_readRange); err == nil {
		readRange = e
	}

	writeMap:=make(map[string]interface{})

	fmt.Printf("准备写数据，当前坐标位置：%s 处理范围： %s-%s\n", excel.GetCellName(this.target.X, this.target.Y), excel.GetCellName(this.target.X, this.target.Y), excel.GetCellName(this.target.Xe, this.target.Ye))

	if len(this.getReadData()) == 0 {
		this.setReadData(append(this.getReadData(), TableMapLineData{value: ""}))
	}
	if len(tableMap) > 0 {
		keyList := []*TableMap{}
		for _, table := range tableMap {
			if table.key {
				keyList = append(keyList, table)
			}
		}

		var line=0
		for _, lineData := range this.getReadData() {
			if readRange > 0 && line >= readRange {
				break
			}
			if len(lineData.lineData) > 0 {
				if len(keyList) > 0 {
					key := ""
					for _, k := range keyList {
						key += lineData.lineData[k.name].value.(string)
						if len(key) > 0 {
							key += ","
						}
					}
					isfind := false
					isAdd := false
					for i := this.target.Ys; i < this.target.Ye; i++ {
						dstkey := ""
						for _, k := range keyList {
							dstkey += this.target.File.GetCellValue(this.target.CurrentSheet, excel.GetCellName(lineData.lineData[k.name].table.dstX, i))
							if len(dstkey) > 0 {
								dstkey += ","
							}
						}
						if dstkey == key {
							isfind = true
							this.target.Y = i
							break
						}
					}
					if !isfind {
						for i := this.target.Ys; i < this.target.Ye; i++ {
							dstkey := ""
							for _, k := range keyList {
								dstkey += this.target.File.GetCellValue(this.target.CurrentSheet, excel.GetCellName(lineData.lineData[k.name].table.dstX, i))
								if len(dstkey) > 0 {
									dstkey += ","
								}
							}
							if dstkey == "" {
								isfind = true
								isAdd = true
								this.target.Y = i
								break
							}
						}
					}
					if !isfind {
						panic("未找到可插入的位置！！！！")
					}
					if isAdd {
						nextIsNull := true
						for _, data := range lineData.lineData {
							if data.table.dstColumn != "" {
								if this.target.File.GetCellValue(this.target.CurrentSheet, excel.GetCellName(data.table.dstX, this.target.Y)) != "" {
									nextIsNull = false
									break
								}
							}
						}
						if !nextIsNull {
							this.target.File.InsertRow(this.target.CurrentSheet, this.target.Y)
						}
					}
				}
				for k:=range writeMap{
					delete(writeMap,k)
				}
				for _, data := range lineData.lineData {
					if data.table.dstColumn != "" {
						if v,ok:=writeMap[data.table.dstColumn];ok{
							if data.table.format==define.KEY_VALUE_float{
								fval, _ := strconv.ParseFloat(data.value.(string), 64)
								fival,_:=strconv.ParseFloat(v.(string),64)
								if fival>fval{
									fmt.Printf("重复写入 %s 的值 %v 太小\n", excel.GetCellName(this.target.X, this.target.Y), fval)
									continue
								}
							}

						}
						writeMap[data.table.dstColumn]=data.value
						this.writeNewValue("", data.value, data.table.format, excel.GetCellName(data.table.dstX, this.target.Y))
					}
				}

			} else {
				val := this.target.File.GetCellValue(this.target.CurrentSheet, excel.GetCellName(this.target.X, this.target.Y))
				if val != "" {
					fmt.Printf("覆盖 %s 的值 %v \n", excel.GetCellName(this.target.X, this.target.Y), val)
				}
				this.writeNewValue("", lineData.value, format, excel.GetCellName(this.target.X, this.target.Y))
			}
			line++
			this.target.Y++
		}

	} else {
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
	}
	this.setReadData([]TableMapLineData{})
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
		var err error
		if wValue, err = strconv.ParseFloat(newval.(string), 64); err != nil {
			wValue = newval
		}
		//value := fmt.Sprintf("%v", newval)
		//dateValue := "2006-01-02T00:00:00+08:00"
		//if fv:= strings.Split(value,"-");len(fv)==2{
		//	//8位日期
		//	dateValue = strings.Replace(dateValue, "2006", "20"+fv[2], -1)
		//	dateValue = strings.Replace(dateValue, "-01-", fmt.Sprintf("-%s-", fv[0]), -1)
		//	dateValue = strings.Replace(dateValue, "02T", fmt.Sprintf("%sT", fv[1]), -1)
		//	t, err := time.Parse(time.RFC3339, dateValue)
		//	if dateValue != "" && err != nil {
		//		panic(err)
		//	}
		//	wValue = t
		//}else{
		//	wValue=value
		//}
		//style, _ = this.target.File.NewStyle(`{"custom_number_format": "yyyy/mm/dd"}`)
		//style, err := xlsx.NewStyle(`{"border":[{"type":"left","color":"0000FF","style":3},{"type":"top","color":"00FF00","style":4},{"type":"bottom","color":"FFFF00","style":5},{"type":"right","color":"FF0000","style":6},{"type":"diagonalDown","color":"A020F0","style":7},{"type":"diagonalUp","color":"A020F0","style":8}]}`)
		style, _ = this.target.File.NewStyle(`{"border":[{"type":"left","color":"000000","style":1},{"type":"top","color":"000000","style":1},{"type":"bottom","color":"000000","style":1},{"type":"right","color":"000000","style":1}],"custom_number_format": "[$-380A]yyyy/mm/dd"}`)
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
					panic(fmt.Sprintf("写数据格式转换错误：%s sheet: %s cell:%s", this.target.File.Path, this.target.CurrentSheet, cellName))
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

	var readSrc bool
	if this.target == this.data[define.KEY_VALUE_src].(*excel.ExcelObject) {
		readSrc = true
		this.setReadData([]TableMapLineData{})
	} else {
		readSrc = false
	}

	fmt.Printf("准备读数据，当前坐标位置：%s，处理范围： %s-%s\n", excel.GetCellName(this.target.X, this.target.Y), excel.GetCellName(this.target.X, this.target.Y), excel.GetCellName(this.target.Xe, this.target.Ye))
	if this.target.Operation == "" {
		this.setReadData(append(this.getReadData(), TableMapLineData{value: this.target.File.GetCellValue(this.target.CurrentSheet, excel.GetCellName(this.target.X, this.target.Y))}))
		fmt.Printf("坐标 %s 读到数据：%v\n", excel.GetCellName(this.target.X, this.target.Y), this.getReadData())

	} else if this.target.Operation == define.KEY_VALUE_right {
		for i := this.target.X; i < this.target.Xe; i++ {
			val := this.target.File.GetCellValue(this.target.CurrentSheet, excel.GetCellName(i, this.target.Y))
			if val == readEndCondition {
				break
			}
			if readRange > 0 && this.target.Xe-i > readRange {
				break
			}
			this.setReadData(append(this.getReadData(), TableMapLineData{value: val}))
		}

	} else if this.target.Operation == define.KEY_VALUE_down {
		line := 0

		for ; this.target.Y < this.target.Ye; this.target.Y++ {
			if readRange > 0 && line >= readRange {
				break
			}
			if len(tableMap) > 0 {
				isbreak := false
				lineData := TableMapLineData{}
				for _, tm := range tableMap {
					targetX := 0
					if readSrc {
						targetX = tm.srcX
					} else {
						targetX = tm.dstX
					}
					val := this.target.File.GetCellValue(this.target.CurrentSheet, excel.GetCellName(targetX, this.target.Y))
					if tm.format == define.KEY_VALUE_date {
						style, _ := this.target.File.NewStyle(`{"border":[{"type":"left","color":"000000","style":1},{"type":"top","color":"000000","style":1},{"type":"bottom","color":"000000","style":1},{"type":"right","color":"000000","style":1}],"custom_number_format": "yyyy/mm/dd"}`)
						this.target.File.SetCellStyle(this.target.CurrentSheet, excel.GetCellName(targetX, this.target.Y), excel.GetCellName(targetX, this.target.Y), style)
						val = this.target.File.GetCellValue(this.target.CurrentSheet, excel.GetCellName(targetX, this.target.Y))
					}
					if val == tm.readEndCondition {
						isbreak = true
						break
					}
					if lineData.lineData == nil {
						lineData.lineData = map[string]TableMapValue{}
					}
					if readSrc && tm.srcColumn == "" && tm.value != "" {
						val = tm.value
					}
					lineData.lineData[tm.name] = TableMapValue{table: tm, srcY: line, dstY: line, value: val}
				}
				if isbreak {
					break
				} else {
					fmt.Printf("读到第 %d 行数据：%v\n", this.target.Y, lineData)
					this.setReadData(append(this.getReadData(), lineData))
				}
			} else {
				val := this.target.File.GetCellValue(this.target.CurrentSheet, excel.GetCellName(this.target.X, this.target.Y))
				if val == readEndCondition {
					break
				}
				if readRange > 0 && line > readRange {
					break
				}
				this.setReadData(append(this.getReadData(), TableMapLineData{value: val}))
				fmt.Printf("读到数据%d行：%v\n", this.target.Y, this.getReadData())
			}
			line++
		}

	}
}

func (this *YgcOfficeProcess) ProcessXEndText(s string) {
	isFindData := false
	if s == key_data {
		s = this.getReadData()[0].value.(string)
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
		s = this.getReadData()[0].value.(string)
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
		s = this.getReadData()[0].value.(string)
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
		s = this.getReadData()[0].value.(string)
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
	fmt.Printf("设置X坐标：%s 条件：%s\n", excel.GetCellName(this.target.X, this.target.Y), s)
}

func (this *YgcOfficeProcess) ProcessYStartText(s string) {
	if s == key_data {
		s = this.getReadData()[0].value.(string)
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
	fmt.Printf("设置Y坐标：%s 条件：%s\n", excel.GetCellName(this.target.X, this.target.Y), s)
}
func (this *YgcOfficeProcess) ProcessXStartText(s string) {
	if s == key_data {
		s = this.getReadData()[0].value.(string)
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
	this.setReadData(append(this.getReadData(), TableMapLineData{value: val}))
	fmt.Printf("在 %s 读到数据：%v\n", pos, this.getReadData())
}
func (this *YgcOfficeProcess) ProcessTableMap(cfg *config.Config, s string) {
	tableMap = []*TableMap{}
	ss, _ := cfg.String(s, define.KEY_OPTION_tableMap)
	ssp := strings.Split(ss, ",")
	var err error
	for _, section := range ssp {
		tableobj := TableMap{}
		tableobj.name = section
		tableobj.srcColumn, _ = cfg.String(section, define.KEY_OPTION_srcColumn)
		tableobj.dstColumn, _ = cfg.String(section, define.KEY_OPTION_dstColumn)
		tableobj.format, _ = cfg.String(section, define.KEY_OPTION_format)
		tableobj.key, _ = cfg.Bool(section, define.KEY_OPTION_key)
		tableobj.value, _ = cfg.String(section, define.KEY_OPTION_value)
		if tableobj.value != "" {
			tableobj.value = findByMap(&this.data, tableobj.value)
		}

		if tableobj.readEndCondition, err = cfg.String(section, define.KEY_OPTION_readEndCondition); err != nil {
			tableobj.readEndCondition = "================================="
		}

		backTarget := this.target
		this.ProcessTarget(define.KEY_VALUE_src)
		this.goPostion(tableobj.srcColumn)
		tableobj.srcX = this.target.X
		this.ProcessTarget(define.KEY_VALUE_dst)
		this.goPostion(tableobj.dstColumn)
		tableobj.dstX = this.target.X
		this.target = backTarget
		tableMap = append(tableMap, &tableobj)
	}
}
func (this *YgcOfficeProcess) goPostion(s string) {
	if s == "" {
		return
	}
	spsrcColumn := strings.Split(s, ",")
	text := findByMap(&this.data, spsrcColumn[0])
	this.ProcessXFindText(text)
	if len(spsrcColumn) > 1 {
		if ax, err := strconv.ParseInt(spsrcColumn[1], 10, 32); err == nil {
			this.target.X += int(ax)
		}
	}
	if len(spsrcColumn) > 2 {
		if ay, err := strconv.ParseInt(spsrcColumn[2], 10, 32); err == nil {
			this.target.Y += int(ay)
		}
	}
}
func (this *YgcOfficeProcess) processSum(cfg *config.Config, section string) {
	ss, _ := cfg.String(section, define.KEY_OPTION_sumSection)
	var wValue float64 = 0
	var wData TableMapLineData
	for _, ld := range this.getReadData() {
		if v, ok := ld.lineData[ss]; ok {
			switch v.table.format {
			//case define.KEY_VALUE_string:
			//	wValue = fmt.Sprintf("%v%v", wValue, v.value)
			case define.KEY_VALUE_float:

				if value, err := strconv.ParseFloat(v.value.(string), 64); err == nil {
					wValue = wValue + value
				} else {
					panic(err)
				}
			default:
				panic("format 未处理的类型 " + v.table.format)
			}
		}
	}
	wData.value = wValue
	this.setReadData([]TableMapLineData{wData})
	fmt.Printf("统计得到数据：%v\n", this.getReadData())

}

type TypeSort []TableMapLineData

var SortField string

func (c TypeSort) Len() int {
	return len(c)
}
func (c TypeSort) Swap(i, j int) {
	c[i], c[j] = c[j], c[i]
}
func (c TypeSort) Less(i, j int) bool {
	a := c[i].lineData[SortField].value.(string)
	b := c[j].lineData[SortField].value.(string)
	if a == "" {
		a = "0"
	}
	if b == "" {
		b = "0"
	}
	fa, _ := strconv.ParseFloat(a, 64)
	fb, _ := strconv.ParseFloat(b, 64)
	return fa > fb
}
func (this *YgcOfficeProcess) processSort(cfg *config.Config, section string) {
	if ss, err := cfg.String(section, define.KEY_OPTION_value); err == nil {
		SortField = ss
		var sortdata TypeSort = this.getReadData()
		sort.Sort(sortdata)
		//数据去重
		for i:=len(sortdata)-1;i>=0;i--{
			for l:=i-1;l>=0;l--{
				valueEQ:=true
				for k,v:=range sortdata[i].lineData{
					if dv,_:= sortdata[l].lineData[k];dv.value!=v.value{
						valueEQ=false
						break
					}
				}
				if valueEQ{
					//fmt.Printf("移除重复数据： %v\n",sortdata[l])
					sortdata=append(sortdata[0:l],sortdata[l+1:]...)
					break
				}
			}
		}
		for i,v:=range sortdata{
			fmt.Printf("排序结果%d=%v\n",i,v)
		}
		this.setReadData(sortdata)
	} else {
		panic("未设置排序字段")
	}
}
func (this *YgcOfficeProcess) processLimt(cfg *config.Config, section string) {
	if cnt, err := cfg.Int(section, define.KEY_OPTION_sumSection); err == nil {
		this.setReadData(this.getReadData()[:cnt])
	} else {
		panic("未设置排序字段")
	}
}
func (this *YgcOfficeProcess) processReset(cfg *config.Config, section string) {
	cur := this.target
	this.ProcessTarget(define.KEY_VALUE_src)
	this.target.X, this.target.Xs, this.target.Xe = 0, 0, 0
	this.target.Y, this.target.Ys, this.target.Ye = 0, 0, 0
	this.ProcessTarget(define.KEY_VALUE_dst)
	this.target.X, this.target.Xs, this.target.Xe = 0, 0, 0
	this.target.Y, this.target.Ys, this.target.Ye = 0, 0, 0
	this.target = cur
}
func (this *YgcOfficeProcess) processFilter(cfg *config.Config, section string) {
	if filterField, err := cfg.String(section, define.KEY_OPTION_key); err == nil {
		if vals,err1:= cfg.String(section, define.KEY_OPTION_value);err1==nil {
			fields := strings.Split(filterField, ",")
			fieldVals:=strings.Split(vals,",")
			data := this.getReadData()

			var filterbool [10]bool
			//数据去重
			for i := len(data) - 1; i >= 0; i-- {
				for l, f := range fields {
					val := data[i].lineData[f].value.(string)
					switch data[i].lineData[f].table.format {
					case define.KEY_VALUE_float:
						fval, _ := strconv.ParseFloat(val, 64)
						fival,_:=strconv.ParseFloat(fieldVals[l],64)

						filterbool[l]= fval>fival
					default:
						panic("过滤不支持的类型")
					}
				}
				iscontinue:=false
				for l:=0;l<len(fields);l++{
					if filterbool[l]{
						iscontinue=true
						break
					}
				}
				if iscontinue{
					continue
				}
				data=append(data[0:i],data[i+1:]...)
			}
			for i, v := range data {
				fmt.Printf("过滤结果%d=%v\n", i, v)
			}
			this.setReadData(data)
		}else{
			panic("未设置过滤值")
		}
	} else {
		panic("未设置过滤字段")
	}
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

func createNewChildProcessObject(cfg *config.Config, section string, src excel.ExcelObject, dst *excel.ExcelObject, target *excel.ExcelObject, readData []TableMapLineData, parentSection string) {
	var data = map[string]interface{}{}
	data[define.KEY_VALUE_src] = &src
	data[define.KEY_VALUE_dst] = dst
	data[define.KEY_VALUE_compny] = excel.GetCompnyNameFromPath(src.File.Path)
	var ygcOfficeProcess = YgcOfficeProcess{}
	ygcOfficeProcess.data = data
	ygcOfficeProcess.setReadData(readData)
	ygcOfficeProcess.target = target
	ygcOfficeProcess.ProcessCommand(cfg, section, parentSection)
}

package main

/**
打开表
1、确定读范围
2、确定读方向
3、确定数据处理方式
4、
 */
import (
	"github.com/larspensjo/config"
	"flag"
	"log"
	"ygcOffice/define"
	"fmt"
	"github.com/Luxurioust/excelize"
	"ygcOffice/process"
	"ygcOffice/foreachDir"
	"strings"
	"time"
)

var (
	configFile = flag.String("configfile", "config.ini", "General configuration file")
)

func main() {
	var dstFile string
	var srcFile string
	var dstxlsx *excelize.File
	var srcxlsx *excelize.File
	var srcList []string
	pdstFile := flag.String("dst", "", "区域合并明细文件路径")
	psrcFile := flag.String("src", "", "分公司文件路径")
	flag.Parse()

	cfg, err := config.ReadDefault(*configFile)
	if err != nil {
		log.Printf("%v,没有发现操作配置文件，已经产生了一个默认操作配置文件config.ini，请查看config.ini文件的配置说明进行配置。\n",err)
		cfg = config.NewDefault()
		cfg.AddSection(define.KEY_SECTION_DEMO)
		cfg.AddOption(define.KEY_SECTION_DEMO,define.KEY_OPTION_dstFile,"汇总文件名")
		cfg.AddOption(define.KEY_SECTION_DEMO,define.KEY_OPTION_srcFile,"公司文件名")
		cfg.AddOption(define.KEY_SECTION_DEMO,define.KEY_OPTION_target,"操作目标，src为原文件，dst为目标文件")
		cfg.AddOption(define.KEY_SECTION_DEMO,define.KEY_OPTION_sheet,"需要操作的源表格，为空不操作")
		cfg.AddOption(define.KEY_SECTION_DEMO,define.KEY_OPTION_xFindText,"查找字符串，按照这个字符串开始定x位，为空不操作")
		cfg.AddOption(define.KEY_SECTION_DEMO,define.KEY_OPTION_yFindText,"查找字符串，按照这个字符串开始定y位，为空不操作")
		cfg.AddOption(define.KEY_SECTION_DEMO,define.KEY_OPTION_operation,"right从左到右，down从上往下，sum或为空则叠加")
		cfg.AddOption(define.KEY_SECTION_DEMO,define.KEY_OPTION_xEndText,"查找字符串，按照这个字符串开始定x结束位，为空不操作")
		cfg.AddOption(define.KEY_SECTION_DEMO,define.KEY_OPTION_yEndText,"查找字符串，按照这个字符串开始定y结束位，为空不操作")
		cfg.AddOption(define.KEY_SECTION_DEMO,define.KEY_OPTION_yAdd,"增加x初始坐标位置")
		cfg.AddOption(define.KEY_SECTION_DEMO,define.KEY_OPTION_yAdd,"增加y初始坐标位置")
		cfg.AddOption(define.KEY_SECTION_DEMO,define.KEY_OPTION_readRange,"设置读写范围")
		cfg.AddOption(define.KEY_SECTION_DEMO,define.KEY_OPTION_readEndCondition,"结束条件，如果为空表示判断到空行，填入字符代表结束的字符")
		cfg.AddOption(define.KEY_SECTION_DEMO,define.KEY_OPTION_format,"写数据的格式，string为字符、float为浮点数、int为整数、date为日期")
		cfg.AddOption(define.KEY_SECTION_DEMO,define.KEY_OPTION_process,"数据处理过程，空为不处理，read为把数据都出来、write为把数据写进去、print把数据打印出来")
		cfg.AddOption(define.KEY_SECTION_DEMO,define.KEY_OPTION_nextSection,"操作节点列表")
		cfg.AddOption(define.KEY_SECTION_DEMO,define.KEY_OPTION_columnMap,"源数据和目标数据的列数据映射，如果不能匹配会自动增加一行")
		cfg.AddOption(define.KEY_SECTION_DEMO,define.KEY_OPTION_hasData,"写数据的时候，如果有数据的处理方法，默认是替换，sum是相加，clr是清除")


		cfg.WriteFile(*configFile,0644, "阳光城Office导入程序配置文件 by waroy\ndemo为配置文件说明，激活配置需要设置main节点")
	} else {
		//for _, val := range cfg.Sections() {
		//	log.Println(val)
		//}
	}

	if !cfg.HasSection(define.KEY_SECTION_main){
		log.Println("缺少main配置！")
		return
	}

	if *pdstFile == "" {
		dstfile,_:= cfg.String(define.KEY_SECTION_main, define.KEY_OPTION_dstFile)
		pdstFile=&dstfile
	}

	if *pdstFile == "" {
		print("选择区域合并明细表文件：")
		fmt.Scan(&dstFile)
	} else {
		dstFile = *pdstFile
	}

	if dstxlsx, err = excelize.OpenFile(dstFile); err != nil {
		panic(err)
		return
	} else {
		println("汇总文件加载成功 ", dstFile)
	}

	if *psrcFile == "" {
		srcfile,_:= cfg.String(define.KEY_SECTION_main, define.KEY_OPTION_srcFile)
		psrcFile=&srcfile
	}

	if *psrcFile=="dir"{
		if flist,err:= foreachDir.ListDir(".",".xlsx");err==nil{
			for _,f:=range flist{
				fs:=strings.Split(f,"-")

				if len(fs)==3 && strings.Index(fs[0], ".\\公司")==0{
					srcList=append(srcList,f)
				}
			}
		}else {
			panic(err)
		}
	}

	if *psrcFile==""{
		print("选择一个子公司文件：")
		fmt.Scan(&srcFile)
	}else{
		srcFile=*psrcFile
	}

	startTime:=time.Now()

	if srcFile!="dir" {
		if srcxlsx, err = excelize.OpenFile(srcFile); err != nil {
			println(err)
			return
		} else {
			println("分公司文件加载成功 ", srcFile)
		}
		process.NewProcess(cfg,define.KEY_SECTION_main,srcxlsx,dstxlsx)
	}else{
		for _,file:=range srcList{
			if srcxlsx, err = excelize.OpenFile(file); err != nil {
				println(err)
				return
			} else {
				println("分公司文件加载成功 ", file)
			}
			process.NewProcess(cfg,define.KEY_SECTION_main,srcxlsx,dstxlsx)
		}
	}


	dstxlsx.Save()

	//stop:= time.NewTimer(time.Second)
	//<- stop.C
	//stop.Stop()
	waitTime:= time.Now().Sub(startTime)
	fmt.Printf("耗时 %s 程序处理完成，按任意键退出……",waitTime)
	var onkey string
	fmt.Scanln(onkey)
}


package catch

import (
	"fmt"
	"runtime"
)



/*
Func 发生错误时执行的程序
param 用于记录执行的参数
 */
func Catch(Func func(...interface{}),param ... interface{}) {
	// 必须要先声明defer，否则不能捕获到panic异常
	if err := recover(); err != nil {
		buf := make([]byte, 1024)
		for {
			n := runtime.Stack(buf, false)
			if n < len(buf) {
				buf= buf[:n]
				break
			}
			buf = make([]byte, 2*len(buf))
		}
		stack :=string(buf)
		//panicindex:= strings.Index(stack,"panic.go:458")
		//stack = stack[panicindex:]
		//errindex:= strings.Index(stack,"\n")
		//stack = stack[errindex:]
		fmt.Printf("错误信息：%v\n args:%v\n %s\n",err,param, stack)
		WriteError(err,param, stack)
		if Func!=nil{
			Func(param...)
		}
	}
}


func WriteError(error interface{}, param interface{},stack string)  {

}


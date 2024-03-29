package Record2Excel

import "fmt"

type GlobalConfig struct {
	PrintLog bool
}

var globalConfig = GlobalConfig{
	PrintLog: false,
}

func log(v ...any) {
	if !globalConfig.PrintLog {
		return
	}
	fmt.Println(v...)
}

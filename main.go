package main

import (
	"flag"
	"fmt"
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/tidwall/gjson"
	"golang.org/x/sys/windows/registry"
	"io"
	"log"
	"net/http"
	"net/url"
	"os"
	"path/filepath"
	"runtime"
	"runtime/debug"
	"strings"
	"sync"
	"syscall"
	"unsafe"
)

const (
	HKEY_CURRENT_USER = 0x80000001
	REG_SZ            = 1

	KEY_WRITE = 0x20006
	KEY_READ  = 0x20019

	SPI_SETDESKWALLPAPER = 0x0014
	SPIF_UPDATEINIFILE   = 0x01
	SPIF_SENDCHANGE      = 0x02
)

var (
	kernel32        = syscall.MustLoadDLL("kernel32.dll")
	getModuleHandle = kernel32.MustFindProc("GetModuleHandleW")

	advapi32      = syscall.MustLoadDLL("Advapi32.dll")
	regOpenKeyEx  = advapi32.MustFindProc("RegOpenKeyExW")
	regSetValueEx = advapi32.MustFindProc("RegSetValueExW")
	regCloseKey   = advapi32.MustFindProc("RegCloseKey")
)

func init() {
	// 开机启动

	// 注册任务事件
}

func main() {
	// 关闭GC
	debug.SetGCPercent(-1)

	// 获取壁纸文件
	image, err := DownloadBingImage()
	if err != nil {
		return
	}

	// 设置壁纸

	// 刷新桌面

	// 替换为你实际的壁纸文件路径
	err = setWallpaper(image)
	if err != nil {
		fmt.Println("Error setting wallpaper:", err)
		return
	}
	// 删除壁纸文件
	err = os.Remove(image)
	if err != nil {
		// 如果发生错误，打印出来
		fmt.Println("删除壁纸文件错误:", err)
	}
	ts := flag.Bool("ts", false, "设置Windows任务计划，可在taskschd.msc中查看")
	flag.Parse()
	if *ts {
		// 创建任务计划
		createSchedule()
	}
}

func DownloadFile(url, fileName string) error {
	// 获取响应头
	resp, err := http.Head(url)
	if err != nil {
		return err
	}
	// 获取文件大小
	size := int(resp.ContentLength)
	/*size, err := strconv.Atoi(resp.Header.Get("Content-Length"))
	  if err != nil {
	      return err
	  }*/
	cd := resp.Header.Get("Content-Disposition") // 可能包含文件名
	log.Println(cd)
	ranges := resp.Header.Get("Accept-Ranges") // bytes 支持范围请求，None 不支持范围请求

	// 创建文件
	//filename := path.Base(url)
	//file, err := os.OpenFile(filename, os.O_CREATE|os.O_WRONLY, 0666)

	file, err := os.Create(fileName)
	if err != nil {
		return err
	}
	defer func(out *os.File) {
		err := out.Close()
		if err != nil {
			log.Println(err)
		}
	}(file)

	if ranges != "bytes" || size <= 10000 { // 当文件过小时单线程下载
		resp, err := http.Get(url)
		if err != nil {
			return err
		}
		defer func(Body io.ReadCloser) {
			err := Body.Close()
			if err != nil {
				log.Println(err)
			}
		}(resp.Body)

		_, err = io.Copy(file, resp.Body)
		if err != nil {
			return err
		}
		return nil
	}

	concurrency := 10 // 并发数
	// 控制并发
	var wg sync.WaitGroup
	wg.Add(concurrency)

	var bg int64 // 起始位置
	var ed int64 // 结束位置
	for i := 0; i < concurrency; i++ {
		bg = int64(i) * int64(size/concurrency)
		ed = bg + int64(size/concurrency) - 1

		go func(idx int, bg, ed int64) {
			defer wg.Done()

			req, _ := http.NewRequest(http.MethodGet, url, nil)
			req.Header.Set("Range", fmt.Sprintf("bytes=%v-%v", bg, ed))

			resp, err := http.DefaultClient.Do(req)
			//client := &http.Client{}
			//resp, err := client.Do(req)
			if err != nil {
				panic(err)
			}
			defer func(Body io.ReadCloser) {
				err := Body.Close()
				if err != nil {
					log.Println(err)
				}
			}(resp.Body)

			_, err = file.Seek(bg, 0)
			if err != nil {
				panic(err)
			}
			_, err = io.Copy(file, resp.Body)
			if err != nil {
				panic(err)
			}
			log.Printf("[%d] Done.", idx)
		}(i, bg, ed)
	}
	wg.Wait()
	return nil
}

// DownloadBingImage 下载必应壁纸
func DownloadBingImage() (string, error) {
	// https://assets.msn.cn/resolver/api/resolve/v3/config/?expType=AppConfig&expInstance=default&apptype=edgeChromium&v=20240202.634
	// BackgroundImageWC/default.properties.cmsImage

	// 获取壁纸图片地址
	const bingAPI = "https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1"
	resp, err := http.Get(bingAPI)
	if err != nil {
		return "", err
	}
	defer func(Body io.ReadCloser) {
		err := Body.Close()
		if err != nil {
			log.Println(err)
		}
	}(resp.Body)
	body, err := io.ReadAll(resp.Body)
	if err != nil {
		return "", err
	}
	// 提取图片地址
	imageUrl := "https://www.bing.com" + gjson.GetBytes(body, "images.0.url").String()
	u, err := url.Parse(imageUrl) // 解析URL
	if err != nil {
		return "", err
	}
	q := u.Query() // 获取查询参数
	id := q.Get("id")
	fmt.Println(id)
	rf := q.Get("rf")
	fmt.Println(rf)

	response, e := http.Get(imageUrl)
	if e != nil {
		return "", e
	}
	defer func(Body io.ReadCloser) {
		err := Body.Close()
		if err != nil {
			log.Println(err)
		}
	}(response.Body)
	// 获取图片格式
	ext := filepath.Ext(rf)
	if ext == "" {
		ext = "jpg"
	}
	// 去掉前缀
	ext = strings.TrimPrefix(ext, ".")
	// 创建图片文件
	file, err := os.Create("bing_wallpaper." + ext)
	if err != nil {
		return "", err
	}
	defer func(file *os.File) {
		err := file.Close()
		if err != nil {
			log.Println(err)
		}
	}(file)

	// 保存图片到文件
	_, err = io.Copy(file, response.Body)
	if err != nil {
		return "", err
	}

	// 返回图片文件路径
	filePath, err := filepath.Abs(file.Name())
	if err != nil {
		log.Fatal(err)
	}
	return filePath, nil
}

// 设置壁纸
func setWallpaper(filePath string) error {
	// 获取系统参数信息函数
	syscallSPI_SETDESKWALLPAPER := syscall.MustLoadDLL("user32.dll").MustFindProc("SystemParametersInfoW")

	// 转换文件路径为UTF16编码的指针
	filePathUTF16Ptr, err := syscall.UTF16PtrFromString(filePath)
	if err != nil {
		return err
	}
	// 调用SystemParametersInfoW函数来设置壁纸
	ret, _, err := syscallSPI_SETDESKWALLPAPER.Call(
		uintptr(SPI_SETDESKWALLPAPER),
		0,
		uintptr(unsafe.Pointer(filePathUTF16Ptr)),
		uintptr(SPIF_UPDATEINIFILE|SPIF_SENDCHANGE),
	)
	/*ret, _, err := systemParametersInfo.Call(
	    uintptr(0x0014),
	    uintptr(0),
	    uintptr(unsafe.Pointer(filePathPtr)),
	    uintptr(0x01|0x02),
	)*/

	if ret == 0 {
		return fmt.Errorf("SystemParametersInfoW call failed with error: %v", err)
	}
	// https://blog.csdn.net/CodyGuo/article/details/73013557
	return nil
}

func setRegistryValue(key syscall.Handle, subKey, valueName, data string) error {
	var pSubKey, pValueName, pData *uint16
	var err error

	pSubKey, err = syscall.UTF16PtrFromString(subKey)
	if err != nil {
		return err
	}
	pValueName, err = syscall.UTF16PtrFromString(valueName)
	if err != nil {
		return err
	}
	pData, err = syscall.UTF16PtrFromString(data)
	if err != nil {
		return err
	}

	_, _, err = regOpenKeyEx.Call(
		uintptr(key),
		uintptr(unsafe.Pointer(pSubKey)),
		0,
		uintptr(KEY_WRITE),
		uintptr(unsafe.Pointer(&key)),
	)
	if err != nil {
		return err
	}
	defer func(regCloseKey *syscall.Proc, a ...uintptr) {
		_, _, err := regCloseKey.Call(a...)
		if err != nil {
			fmt.Println("Error closing registry key:", err)
		}
	}(regCloseKey, uintptr(key))

	_, _, err = regSetValueEx.Call(
		uintptr(key),
		uintptr(unsafe.Pointer(pValueName)),
		0,
		uintptr(REG_SZ),
		uintptr(unsafe.Pointer(pData)),
		uintptr(len(data)),
	)
	if err != nil {
		return err
	}

	return nil
}

// 开机启动
func addToStartup(appName, appPath string) error {
	k, err := registry.OpenKey(registry.CURRENT_USER, `Software\Microsoft\Windows\CurrentVersion\Run`, registry.SET_VALUE)
	if err != nil {
		return err
	}
	defer func(k registry.Key) {
		err := k.Close()
		if err != nil {
			fmt.Println("Error closing registry key:", err)
		}
	}(k)
	/*exePath, err := os.Executable()
	  if err != nil {
	      fmt.Println("Error getting executable path:", err)
	      return
	  }

	  // Ensure the path is absolute.
	  exePath, err = filepath.Abs(exePath)
	  if err != nil {
	      fmt.Println("Error getting absolute executable path:", err)
	      return
	  }*/
	err = k.SetStringValue(appName, `"`+appPath+`"`)
	if err != nil {
		return err
	}
	return nil
}

// 创建任务计划: https://docs.microsoft.com/zh-cn/windows/win32/taskschd/task-scheduler-start-page
// see https://github.com/capnspacehook/taskmaster/blob/master/manage.go
// 解锁、启动、登录等事件触发任务计划 taskschd.msc
func createSchedule() {
	sysType := runtime.GOOS
	if sysType == "windows" {
		err := ole.CoInitialize(0)
		if err != nil {
			return
		}
		defer ole.CoUninitialize()

		unknown, _ := oleutil.CreateObject("Schedule.Service")
		defer unknown.Release()

		scheduler, _ := unknown.QueryInterface(ole.IID_IDispatch)
		defer scheduler.Release()

		_, err = oleutil.CallMethod(scheduler, "Connect")
		if err != nil {
			return
		}

		folder, _ := oleutil.CallMethod(scheduler, "GetFolder", "\\")
		defer func(folder *ole.VARIANT) {
			err := folder.Clear()
			if err != nil {

			}
		}(folder)

		tasks, _ := oleutil.CallMethod(folder.ToIDispatch(), "GetTasks", 1)
		defer func(tasks *ole.VARIANT) {
			err := tasks.Clear()
			if err != nil {

			}
		}(tasks)
		err = oleutil.ForEach(tasks.ToIDispatch(), func(v *ole.VARIANT) error {
			task := v.ToIDispatch()
			//name := oleutil.MustGetProperty(task, "Name").ToString()
			name, err := oleutil.GetProperty(task, "Name")
			if err != nil {
				return nil
			}
			println(name.ToString())
			return nil
		})
		// https://github.com/capnspacehook/taskmaster/blob/master/fill.go
		task_definition := oleutil.MustCallMethod(folder.ToIDispatch(), "NewTask", 0).ToIDispatch()
		defer task_definition.Release()
		triggers := oleutil.MustGetProperty(task_definition, "Triggers").ToIDispatch()
		defer triggers.Release()
		registration_info := oleutil.MustGetProperty(task_definition, "RegistrationInfo").ToIDispatch()
		defer registration_info.Release()
		actions := oleutil.MustGetProperty(task_definition, "Actions").ToIDispatch()
		defer actions.Release()
		principal := oleutil.MustGetProperty(task_definition, "Principal").ToIDispatch()
		defer principal.Release()
		settings := oleutil.MustGetProperty(task_definition, "Settings").ToIDispatch()
		defer settings.Release()

		/*repetition := oleutil.MustGetProperty(triggers, "Repetition").ToIDispatch()
		  defer repetition.Release()
		  oleutil.MustPutProperty(repetition, "Duration", "")
		  oleutil.MustPutProperty(repetition, "Interval", "")
		  oleutil.MustPutProperty(repetition, "StopAtDurationEnd", true)*/

		/*//trigger0 := triggers.MustQueryInterface(ole.NewGUID("{d45b0167-9653-4eef-b94f-0732ca7af251}"))
		        trigger0 := oleutil.MustCallMethod(triggers, "Create", uint(0)).ToIDispatch()
		        defer trigger0.Release()
		        oleutil.MustPutProperty(trigger0, "Id", "")
		        oleutil.MustPutProperty(trigger0, "Enabled", true)
		        //oleutil.MustPutProperty(trigger0, "EndBoundary", "")
		        //oleutil.MustPutProperty(trigger0, "ExecutionTimeLimit", "")
		        oleutil.MustPutProperty(trigger0, "Id", "")
		        oleutil.MustPutProperty(trigger0, "Subscription", `<QueryList>
		    <Query Id='0'>
		        <Select Path='System'>
		            *[System[Provider[@Name='Microsoft-Windows-Power-Troubleshooter'] and EventID=1]]
		        </Select>
		    </Query>
		    <Query Id='1'>
		        <Select Path='System'>
		            *[System/Level=2]
		        </Select>
		    </Query>
		</QueryList>`)*/

		/*//trigger1 := triggers.MustQueryInterface(ole.NewGUID("{b45747e0-eba7-4276-9f29-85c5bb300006}"))
		  trigger1 := oleutil.MustCallMethod(triggers, "Create", uint(1)).ToIDispatch()
		  defer trigger1.Release()
		    oleutil.MustPutProperty(trigger1, "Id", "bing_wallpaper_time_trigger")
		    oleutil.MustPutProperty(trigger1, "Enabled", true)*/

		/*//trigger2 := triggers.MustQueryInterface(ole.NewGUID("{126c5cd8-b288-41d5-8dbf-e491446adc5c}"))
		  trigger2 := oleutil.MustCallMethod(triggers, "Create", uint(2)).ToIDispatch()
		  defer trigger2.Release()
		  oleutil.MustPutProperty(trigger2, "Id", "bing_wallpaper_daily_trigger")
		  oleutil.MustPutProperty(trigger2, "Enabled", true)
		  oleutil.MustPutProperty(trigger2, "DaysInterval", 1)*/

		/*//trigger3 := triggers.MustQueryInterface(ole.NewGUID("{5038fc98-82ff-436d-8728-a512a57c9dc1}"))
		  trigger3 := oleutil.MustCallMethod(triggers, "Create", uint(3)).ToIDispatch()
		  defer trigger3.Release()
		  oleutil.MustPutProperty(trigger3, "Id", "bing_wallpaper_weekly_trigger")
		  oleutil.MustPutProperty(trigger3, "Enabled", true)*/

		/*//trigger4 := triggers.MustQueryInterface(ole.NewGUID("{97c45ef1-6b02-4a1a-9c0e-1ebfba1500ac}"))
		  trigger4 := oleutil.MustCallMethod(triggers, "Create", uint(4)).ToIDispatch()
		  defer trigger4.Release()
		  oleutil.MustPutProperty(trigger4, "Id", "bing_wallpaper_monthly_trigger")
		  oleutil.MustPutProperty(trigger4, "Enabled", true)*/

		/*//trigger5 := triggers.MustQueryInterface(ole.NewGUID("{77d025a3-90fa-43aa-b52e-cda5499b946a}"))
		  trigger5 := oleutil.MustCallMethod(triggers, "Create", uint(5)).ToIDispatch()
		  defer trigger5.Release()
		  oleutil.MustPutProperty(trigger5, "Id", "bing_wallpaper_monthlydow_trigger")
		  oleutil.MustPutProperty(trigger5, "Enabled", true)*/

		// 创建闲置触发，在发生空闲情况时启动任务的触发器
		/*//trigger6 := triggers.MustQueryInterface(ole.NewGUID("{d537d2b0-9fb3-4d34-9739-1ff5ce7b1ef3}"))
		  trigger6 := oleutil.MustCallMethod(triggers, "Create", uint(6)).ToIDispatch()
		  defer trigger6.Release()
		  oleutil.MustPutProperty(trigger6, "Id", "bing_wallpaper_idle_trigger")
		  oleutil.MustPutProperty(trigger6, "Enabled", true)*/

		// 创建注册触发器
		//trigger7 := triggers.MustQueryInterface(ole.NewGUID("{4c8fec3a-c218-4e0c-b23d-629024db91a2}"))
		trigger7 := oleutil.MustCallMethod(triggers, "Create", uint(7)).ToIDispatch()
		defer trigger7.Release()
		oleutil.MustPutProperty(trigger7, "Id", "bing_wallpaper_registration_trigger")
		oleutil.MustPutProperty(trigger7, "Enabled", true)

		// 创建启动触发器
		//trigger8 := triggers.MustQueryInterface(ole.NewGUID("{2a9c35da-d357-41f4-bbc1-207ac1b1f3cb}"))
		trigger8 := oleutil.MustCallMethod(triggers, "Create", uint(8)).ToIDispatch()
		defer trigger8.Release()
		oleutil.MustPutProperty(trigger8, "Id", "bing_wallpaper_boot_trigger")
		oleutil.MustPutProperty(trigger8, "Enabled", true)

		// 创建登录触发器
		//trigger9 := triggers.MustQueryInterface(ole.NewGUID("{72dade38-fae4-4b3e-baf4-5d009af02b1c}"))
		trigger9 := oleutil.MustCallMethod(triggers, "Create", uint(9)).ToIDispatch()
		defer trigger9.Release()
		oleutil.MustPutProperty(trigger9, "Id", "bing_wallpaper_logon_trigger")
		oleutil.MustPutProperty(trigger9, "Enabled", true)

		// 用于触发控制台连接或断开连接，远程连接或断开连接或工作站锁定或解锁通知的任务。
		//trigger11 := triggers.MustQueryInterface(ole.NewGUID("{754da71b-4385-4475-9dd9-598294fa3641}"))
		trigger11 := oleutil.MustCallMethod(triggers, "Create", uint(11)).ToIDispatch()
		defer trigger11.Release()
		oleutil.MustPutProperty(trigger11, "Id", "bing_wallpaper_ssc_trigger")
		oleutil.MustPutProperty(trigger11, "Enabled", true)
		// 获取或设置将触发任务启动的终端服务器会话更改的类型：7锁定；8解锁
		oleutil.MustPutProperty(trigger11, "StateChange", uint(8))

		// 自定义
		/*trigger12 := oleutil.MustCallMethod(triggers, "Create", uint(12)).ToIDispatch()
		  defer trigger12.Release()
		  oleutil.MustPutProperty(trigger12, "Id", "")
		  oleutil.MustPutProperty(trigger12, "Enabled", true)*/

		// 设置任务的注册信息
		oleutil.MustGetProperty(registration_info, "Author", "bajins")
		//oleutil.MustPutProperty(registration_info, "Date", "")
		oleutil.MustGetProperty(registration_info, "Description", "设置Bing桌面壁纸")
		//oleutil.MustPutProperty(registration_info, "Documentation", "")
		//oleutil.MustPutProperty(registration_info, "SecurityDescriptor", "")
		//oleutil.MustPutProperty(registration_info, "Source", "")
		//oleutil.MustPutProperty(registration_info, "URI", "")
		//oleutil.MustPutProperty(registration_info, "Version", "")

		// 创建任务的操作
		var context string
		oleutil.MustPutProperty(actions, "Context", context)
		//action := actions.MustQueryInterface(ole.NewGUID("{4c3d624d-fd6b-49a3-b9b7-09cb3cd3f047}"))
		action := oleutil.MustCallMethod(actions, "Create", uint(ole.TKIND_DISPATCH)).ToIDispatch()
		defer action.Release()
		oleutil.MustPutProperty(action, "Id", "set_bing_wallpaper")
		// os.Hostname()
		/*currentUser, err := user.Current()
		  if err != nil {
		      return
		  }
		  fmt.Println(strings.Split(currentUser.Username, `\`)[1])*/
		path, _ := os.Executable()
		_, exec := filepath.Split(path)
		oleutil.MustPutProperty(action, "Path", exec)
		oleutil.MustPutProperty(action, "WorkingDirectory", "")
		oleutil.MustPutProperty(action, "Arguments", "")

		//
		//oleutil.MustPutProperty(principal, "DisplayName", "")
		//oleutil.MustPutProperty(principal, "GroupId", "")
		//oleutil.MustPutProperty(principal, "Id", "")
		oleutil.MustPutProperty(principal, "LogonType", uint(3))
		oleutil.MustPutProperty(principal, "RunLevel", uint(1))
		//oleutil.MustPutProperty(principal, "UserId", "")

		//
		oleutil.MustPutProperty(settings, "Enabled", true)
		oleutil.MustPutProperty(settings, "Hidden", true)
		oleutil.MustPutProperty(settings, "RunOnlyIfIdle", false)
		//oleutil.MustPutProperty(settings, "AllowDemandStart", false)
		//oleutil.MustPutProperty(settings, "AllowHardTerminate", false)
		//oleutil.MustPutProperty(settings, "Compatibility", uint(0))
		//oleutil.MustPutProperty(settings, "DeleteExpiredTaskAfter", false)
		//oleutil.MustPutProperty(settings, "DisallowStartIfOnBatteries", false)
		//oleutil.MustPutProperty(settings, "ExecutionTimeLimit", "")

		/*idlesettingsObj := oleutil.MustGetProperty(settings, "IdleSettings").ToIDispatch()
		  defer idlesettingsObj.Release()
		  oleutil.MustPutProperty(idlesettingsObj, "IdleDuration", "")
		  oleutil.MustPutProperty(idlesettingsObj, "RestartOnIdle", true)
		  oleutil.MustPutProperty(idlesettingsObj, "StopOnIdleEnd", true)
		  oleutil.MustPutProperty(idlesettingsObj, "WaitTimeout", "")*/

		//oleutil.MustPutProperty(settings, "MultipleInstances", uint(0))

		/*networksettingsObj := oleutil.MustGetProperty(settings, "NetworkSettings").ToIDispatch()
		  defer networksettingsObj.Release()
		  oleutil.MustPutProperty(networksettingsObj, "Id", "")
		  oleutil.MustPutProperty(networksettingsObj, "Name", "")*/

		//oleutil.MustPutProperty(settings, "Priority", uint(0))
		//oleutil.MustPutProperty(settings, "RestartCount", uint(0))
		//oleutil.MustPutProperty(settings, "RestartInterval", "")
		//oleutil.MustPutProperty(settings, "RunOnlyIfIdle", true)
		//oleutil.MustPutProperty(settings, "RunOnlyIfNetworkAvailable", true)
		//oleutil.MustPutProperty(settings, "StartWhenAvailable", true)
		//oleutil.MustPutProperty(settings, "StopIfGoingOnBatteries", true)
		//oleutil.MustPutProperty(settings, "WakeToRun", true)

		// 设置任务的注册信息
		oleutil.MustCallMethod(folder.ToIDispatch(), "RegisterTaskDefinition", "SetBingWallpaper", task_definition,
			6, nil, nil, 3)
	} else {
		fmt.Println("111")
	}
}

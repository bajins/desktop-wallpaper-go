package main

import (
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
	"os/user"
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
// 解锁、启动、登录等事件触发任务计划
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
		println(os.Hostname())
		currentUser, err := user.Current()
		if err != nil {
			return
		}
		println(strings.Split(currentUser.Username, `\`)[1])
		path, _ := os.Executable()
		_, exec := filepath.Split(path)
		println(exec)
	} else {
		println("111")
	}
}

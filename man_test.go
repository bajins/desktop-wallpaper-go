package main

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"testing"
)

// https://zhuanlan.zhihu.com/p/606613732
// https://github.com/capnspacehook/taskmaster
func TestT0(t *testing.T) {
	err := ole.CoInitialize(0)
	if err != nil {
		code := err.(*ole.OleError).Code()
		if code != ole.S_OK && code != 0x00000001 {
			return
		}
	}
	schedClassID, err := ole.ClassIDFrom("Schedule.Service")
	if err != nil {
		ole.CoUninitialize()
		return
	}
	taskSchedulerObj, err := ole.CreateInstance(schedClassID, nil)
	if err != nil {
		ole.CoUninitialize()
		return
	}
	if taskSchedulerObj == nil {
		ole.CoUninitialize()
		return
	}
	defer taskSchedulerObj.Release()

	tskSchdlr := taskSchedulerObj.MustQueryInterface(ole.IID_IDispatch)

	_, err = oleutil.CallMethod(tskSchdlr, "Connect")
	if err != nil {
		return
	}
	res, err := oleutil.CallMethod(tskSchdlr, "GetFolder", `\`)
	if err != nil {
		return
	}
	//rootFolderObj := res.ToIDispatch()

	res, err = oleutil.CallMethod(tskSchdlr, "GetRunningTasks", uint(1))
	if err != nil {
		return
	}
	runningTasksObj := res.ToIDispatch()
	defer runningTasksObj.Release()
	err = oleutil.ForEach(runningTasksObj, func(v *ole.VARIANT) error {
		task := v.ToIDispatch()

		runningTask, err := oleutil.GetProperty(task, "EnginePid")
		if err != nil {
			return nil
		}
		println(runningTask)
		return nil
	})
}
func TestT1(t *testing.T) {
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
	res, err := oleutil.CallMethod(scheduler, "GetFolder", `\`)
	if err != nil {
		return
	}
	//rootFolderObj := res.ToIDispatch()

	res, err = oleutil.CallMethod(scheduler, "GetRunningTasks", uint(1))
	if err != nil {
		return
	}
	runningTasksObj := res.ToIDispatch()
	defer runningTasksObj.Release()
	err = oleutil.ForEach(runningTasksObj, func(v *ole.VARIANT) error {
		task := v.ToIDispatch()

		runningTask, err := oleutil.GetProperty(task, "EnginePid")
		if err != nil {
			return nil
		}
		println(runningTask)
		return nil
	})
}

func TestT2(t *testing.T) {
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
	res, err := oleutil.CallMethod(scheduler, "GetFolder", `\`)
	if err != nil {
		return
	}
	//rootFolderObj := res.ToIDispatch()
	tasks, _ := oleutil.CallMethod(res.ToIDispatch(), "GetTasks", 1)
	defer tasks.Clear()

	runningTasksObj := tasks.ToIDispatch()
	//defer runningTasksObj.Release()

	err = oleutil.ForEach(runningTasksObj, func(v *ole.VARIANT) error {
		task := v.ToIDispatch()

		//name := oleutil.MustGetProperty(task, "Name").ToString()
		runningTask, err := oleutil.GetProperty(task, "Name")
		if err != nil {
			return nil
		}
		println(runningTask.ToString())
		return nil
	})
}

func TestT3(t *testing.T) {
	ole.CoInitializeEx(0, ole.COINIT_MULTITHREADED)
	defer ole.CoUninitialize()

	scheduler, _ := oleutil.CreateObject("Schedule.Service")

	task, _ := scheduler.QueryInterface(ole.IID_IDispatch)
	defer task.Release()

	registrationInfo, _ := oleutil.CallMethod(task, "NewTask", 0)

	// 设置任务信息
	oleutil.PutProperty(registrationInfo.ToIDispatch(), "RegistrationInfo", "Author", "Me")
	oleutil.PutProperty(registrationInfo.ToIDispatch(), "RegistrationInfo", "Description", "My task")

	// 创建触发器
	trigger, _ := oleutil.CallMethod(task, "NewTrigger", 0)
	// 设置触发器参数
	_, _ = oleutil.CallMethod(registrationInfo.ToIDispatch(), "RegisterTaskDefinition", "MyTask", task, 6, "", nil, trigger, nil,
		nil, 3)
}

func TestT4(t *testing.T) {
	getEdgeChromiumImageUrl()
}

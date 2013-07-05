package main

import (
	"time"

	"github.com/mattn/go-ole"
	"github.com/mattn/go-ole/oleutil"
)

func main() {
	ole.CoInitialize(0)

	unknown, err := oleutil.GetActiveObject("Outlook.Application")
	created := false
	if err != nil { // not active, create new
		if unknown, err = oleutil.CreateObject("Outlook.Application"); err != nil {
			panic(err) // could not find outlook
		}
		time.Sleep(5 * time.Second) // wait for outlook to start
		created = true
	}

	outlook, _ := unknown.QueryInterface(ole.IID_IDispatch)
	namespace := oleutil.MustCallMethod(outlook, "GetNamespace", "MAPI").ToIDispatch()

	olFolderCalendar := 9
	calendar := oleutil.MustCallMethod(namespace, "GetDefaultFolder", olFolderCalendar).ToIDispatch()
	items := oleutil.MustGetProperty(calendar, "Items").ToIDispatch()

	// items collection should include recurrence patterns
	oleutil.MustPutProperty(items, "IncludeRecurrences", "True")
	oleutil.MustCallMethod(items, "Sort", "[Start]")

	// restrict date range
	layout := "02 Jan 2006"
	start := time.Now().Format(layout)
	end := time.Now().AddDate(0, 0, 1).Format(layout)
	// reversed start/end to filter appointments that overlap this timespan
	restriction := "[End] >= '" + start + "' AND [Start] <= '" + end + "'"

	appointment := oleutil.MustCallMethod(items, "Find", restriction).ToIDispatch()
	for appointment != nil {
		startTime := FloatToTime(oleutil.MustGetProperty(appointment, "Start").ToFloat())
		subject := oleutil.MustGetProperty(appointment, "Subject").ToString()
		println(startTime.Format(time.RFC822) + " - " + subject)

		appointment = oleutil.MustCallMethod(items, "FindNext").ToIDispatch()
	}

	// cleanup
	if created {
		oleutil.CallMethod(outlook, "Quit")
	}
	outlook.Release()
	ole.CoUninitialize()
}

func FloatToTime(v float64) time.Time {
	zero := time.Date(1899, 12, 30, 0, 0, 0, 0, time.Local)
	vd := time.Duration(int64(v * 24 * 60 * 60 * 1000000000))

	return zero.Add(vd)
}

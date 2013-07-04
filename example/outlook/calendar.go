package main

import (
	"log"
	"time"

	"github.com/mattn/go-ole"
	"github.com/mattn/go-ole/oleutil"
)

func main() {
	ole.CoInitialize(0)

	unknown, err := oleutil.GetActiveObject("Outlook.Application")
	created := false
	if err != nil {
		// not active, create new
		if unknown, err = oleutil.CreateObject("Outlook.Application"); err != nil {
			log.Fatal("GetActiveObject/CreatObject:", err)
		}

		created = true
	}

	outlook, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		log.Fatal("QueryInterface:", err)
	}

	// get mapi namespace
	namespaceVariant, err := oleutil.CallMethod(outlook, "GetNamespace", "MAPI")
	if err != nil {
		log.Fatal("GetNamespace:", err)
	}
	namespace := namespaceVariant.ToIDispatch()

	// get calendar folder
	olFolderCalendar := 9
	calendarVariant, err := oleutil.CallMethod(namespace, "GetDefaultFolder", olFolderCalendar)
	if err != nil {
		log.Fatal("GetDefaultFolder:", err)
	}
	calendar := calendarVariant.ToIDispatch()

	// get appointments
	itemsVariant, err := oleutil.GetProperty(calendar, "Items")
	if err != nil {
		log.Fatal("GetProperty Items:", err)
	}
	items := itemsVariant.ToIDispatch()

	oleutil.MustPutProperty(items, "IncludeRecurrences", "True")
	oleutil.MustCallMethod(items, "Sort", "[Start]")

	// restrict date range
	// TODO: Differences in regional date formats?
	layout := "2006/02/01"
	start := time.Now().Format(layout)
	end := time.Now().AddDate(0, 0, 1).Format(layout)
	restriction := "[Start] >= '" + start + "' AND [End] <= '" + end + "'"

	appointment := oleutil.MustCallMethod(items, "Find", restriction).ToIDispatch()
	for appointment != nil {
		startTime := FloatToTime(oleutil.MustGetProperty(appointment, "Start").ToFloat())
		subject := oleutil.MustGetProperty(appointment, "Subject").ToString()

		log.Println(startTime, subject)

		appointment = oleutil.MustCallMethod(items, "FindNext").ToIDispatch()
	}

	log.Println("-- no more appointments --")

	// TODO: is this necessary?
	items.Release()
	calendar.Release()
	namespace.Release()

	// close, if it has been created
	if created {
		oleutil.CallMethod(outlook, "Quit")
	}
	outlook.Release()
	ole.CoUninitialize()
}

func FloatToTime(v float64) time.Time {
	// TODO: timezones?
	zero := time.Date(1899, 12, 30, 0, 0, 0, 0, time.Local)
	vd := time.Duration(int64(v * 24 * 60 * 60 * 1000000000))

	return zero.Add(vd)
}

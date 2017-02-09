// +build windows

package wmi

import (
	"fmt"
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"reflect"
	"runtime"
)

// SWbemServices is used to access wmi. See https://msdn.microsoft.com/en-us/library/aa393719(v=vs.85).aspx
type SWbemServices struct {
	//TODO: track namespace. Not sure if we can re connect to a different namespace using the same instance
	cWMIClient            *Client //This could also be an embedded struct, but then we would need to branch on Client vs SWbemServices in the Query method
	sWbemLocatorIUnknown  *ole.IUnknown
	sWbemLocatorIDispatch *ole.IDispatch
}

// InitializeSWbemServices will return a new SWbemServices object that can be used to query WMI
func InitializeSWbemServices(c *Client, connectServerArgs ...interface{}) (*SWbemServices, error) {
	s := new(SWbemServices)
	s.cWMIClient = c

	//Do we need to force SWbemServices to run on a different thread prior to locking?
	runtime.LockOSThread()

	err := ole.CoInitializeEx(0, ole.COINIT_MULTITHREADED)
	if err != nil {
		oleCode := err.(*ole.OleError).Code()
		if oleCode != ole.S_OK && oleCode != S_FALSE {
			return nil, err
		}
	}

	unknown, err := oleutil.CreateObject("WbemScripting.SWbemLocator")
	if err != nil {
		return nil, err
	} else if unknown == nil {
		return nil, ErrNilCreateObject
	}
	s.sWbemLocatorIUnknown = unknown

	dispatch, err := s.sWbemLocatorIUnknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return nil, err
	}
	s.sWbemLocatorIDispatch = dispatch

	// we can't do the ConnectServer call here unless we find a way to track and re-init the connectServerArgs

	return s, nil
}

// Close will clear and release all of the SWbemServices resources
func (s *SWbemServices) Close() error {
	s.sWbemLocatorIDispatch.Release()
	s.sWbemLocatorIUnknown.Release()
	ole.CoUninitialize()
	runtime.UnlockOSThread()
	return nil
}

// Query runs the WQL query and appends the values to dst.
//
// dst must have type *[]S or *[]*S, for some struct type S. Fields selected in
// the query must have the same name in dst. Supported types are all signed and
// unsigned integers, time.Time, string, bool, or a pointer to one of those.
// Array types are not supported.
//
// By default, the local machine and default namespace are used. These can be
// changed using connectServerArgs. See
// http://msdn.microsoft.com/en-us/library/aa393720.aspx for details.
func (s *SWbemServices) Query(query string, dst interface{}, connectServerArgs ...interface{}) error {
	if s.sWbemLocatorIDispatch == nil {
		return fmt.Errorf("SWbemServices is not Initialized")
	}
	wmi := s.sWbemLocatorIDispatch //Should just rename in the code, but this will help as we break things apart

	dv := reflect.ValueOf(dst)
	if dv.Kind() != reflect.Ptr || dv.IsNil() {
		return ErrInvalidEntityType
	}
	dv = dv.Elem()
	mat, elemType := checkMultiArg(dv)
	if mat == multiArgTypeInvalid {
		return ErrInvalidEntityType
	}

	// service is a SWbemServices
	serviceRaw, err := oleutil.CallMethod(wmi, "ConnectServer", connectServerArgs...)
	if err != nil {
		return err
	}
	service := serviceRaw.ToIDispatch()
	defer serviceRaw.Clear()

	// result is a SWBemObjectSet
	resultRaw, err := oleutil.CallMethod(service, "ExecQuery", query)
	if err != nil {
		return err
	}
	result := resultRaw.ToIDispatch()
	defer resultRaw.Clear()

	count, err := oleInt64(result, "Count")
	if err != nil {
		return err
	}

	enumProperty, err := result.GetProperty("_NewEnum")
	if err != nil {
		return err
	}
	defer enumProperty.Clear()

	enum, err := enumProperty.ToIUnknown().IEnumVARIANT(ole.IID_IEnumVariant)
	if err != nil {
		return err
	}
	if enum == nil {
		return fmt.Errorf("can't get IEnumVARIANT, enum is nil")
	}
	defer enum.Release()

	// Initialize a slice with Count capacity
	dv.Set(reflect.MakeSlice(dv.Type(), 0, int(count)))

	var errFieldMismatch error
	for itemRaw, length, err := enum.Next(1); length > 0; itemRaw, length, err = enum.Next(1) {
		if err != nil {
			return err
		}

		err := func() error {
			// item is a SWbemObject, but really a Win32_Process
			item := itemRaw.ToIDispatch()
			defer item.Release()

			ev := reflect.New(elemType)
			if err = s.cWMIClient.loadEntity(ev.Interface(), item); err != nil {
				if _, ok := err.(*ErrFieldMismatch); ok {
					// We continue loading entities even in the face of field mismatch errors.
					// If we encounter any other error, that other error is returned. Otherwise,
					// an ErrFieldMismatch is returned.
					errFieldMismatch = err
				} else {
					return err
				}
			}
			if mat != multiArgTypeStructPtr {
				ev = ev.Elem()
			}
			dv.Set(reflect.Append(dv, ev))
			return nil
		}()
		if err != nil {
			return err
		}
	}
	return errFieldMismatch
}

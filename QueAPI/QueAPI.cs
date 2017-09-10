using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace QueAPI
{
    using QuePointHandle = System.UInt32;
    using QueSymbolT16 = System.UInt16;
    using QueGarminTimeT32 = System.UInt32;
    using SYSTEMTIME = System.IntPtr;
    using HWND = System.IntPtr;

    public class UnableToConnect : Exception
    {
        public UnableToConnect(string message) : base(message)
        {
        }
    }

    public class QueAPI : IDisposable
    {
        public const int QUE_ADDR_BUF_LENGTH = 50;
        public const int QUE_POST_BUF_LENGTH = 11;

        #region Enums
        private enum QueErrT16 : UInt16
        {
            None = 0, //Success

            //Unsure of the below values

            NotOpen, //Attempted to close the library without opening it first. 
            BadArg, //Invalid parameter passed.
            Memory, //Out of memory.
            NoData, //No data available.
            AlreadyOpen, //The library is already open.
            InvalidVersion, //The library is an incompatible version.
            Comm, // There was an error communicating with the API.
            CmndUnavail, //The command is unavailable.
            StillOpen, //Library is still open.
            Fail, //General failure.
            Cancel //Action cancelled by user.
        };

        //Unsure of these values. Used in GPSSatDataType
        private enum gpsSat : byte
        {
            EphMask = 0x01, //Ephemeris: 0 = no ephemeris, 1 = has ephemeris.
            DifMask = 0x02, //Differential: 0 = no differential correction, 1 = differential correction.
            UsedMask = 0x04, //Used in solution: 0 = no, 1 = yes.
            RisingMask = 0x08 //Satellite rising: 0 = no, 1 = yes.
        }

        //Reason for the sysNotifyGPSDataEvent event
        //Unsure of these values
        public enum QueNotificationT8 : Byte
        {
            LocationChange, // The GPS position has changed.
            StatusChange, // The GPS status has changed.
            LostFix, // The quality of the GPS position computation has become less than two dimensional.
            SatDataChange, // The GPS satellite data has changed.
            ModeChange, // The GPS mode has changed.
            Event, // An generic event has occurred (i.e.sunrise / set, etc.)
            CPOPositionChange, //The GPS CPO position data has been updated.
            SatelliteInstChange, // The GPS CPO satellite data has been updated.
            NavigationEvent //The navigation status has changed. (Added in version 1.50)
        };

        private enum GPSFixT8 : Byte
        {
            Unusable = 0, //GPS failed integrity check.
            Invalid = 1, //GPS is invalid or unavailable.
            TwoD = 2, //Two dimensional position.
            ThreeD = 3, //Three dimensional position.
            TwoDDiff = 4, //Two dimensional differential position. 
            ThreeDDiff = 5 //Three dimensional differential position.
        };

        private enum GPSModeT8 : Byte
        {
            Off = 0, //GPS is off
            Normal = 1, //Continuous satellite tracking or attempting to track satellites.
            BatSaver = 2, //Periodic satellite tracking to conserve battery power (only on iQue).
            Sim = 3, //Simulated GPS information (may be same as gpsModeOff).
            External = 4 //External source of GPS information (only on iQue).
        };

        private enum QueRouteSortT8 : Byte
        {
            None = 0, //Do not apply any sort to the points.
            All = 1, //Sort all of the points.
            IgnoreDest = 3, //Sort all of the points except for the final destination
            IgnoreStart = 5, //Sort all of the points except for the start point.
            IgnoreStartAndDest = 7 //Sort all of the points except for the start point and the final destination.
        };

        public enum QueAppT8 : Byte
        {
            Map, //Default to main map page
            WhereTo, //Default to main search page.
            Gps, //Default to GPS status page
            Turns, //Default to turns list page.
            Trip, //Default to trip computer page.
            Settings, //Default to general settings page.
            GpsSettings, //Default to GPS settings page
            MarkWaypoint, //Default to waypoint marking page.
            Menu, //Default to main menu page (which contains Where To?, Main Map buttons, etc).
            LaunchBackground, //Launch in the background without showing any user interface.
            CloseBackground, //Close the application running in the background with no user interface.  Must be called exactly once for every use of queAppLaunchBackground.
            CloseBackgroundDelay //Close the application running in the background with no user interface but delay first so that if the user immediately does a queAppLaunchBackground, Que will respond quickly.  Must be called exactly once for every use of queAppLaunchBackground.
        };

        private enum QueRouteStatusT8 : Byte
        {
            None = 0, //There is no route currently active.
            Active, //There is a route being actively navigated
            OffRoute, //There is an active route without turn by turn guidance
            Arrived,  //The active route’s destination has been reached.
            Calculating, //A route is being calculated
            Canceled, //The route calculation has been canceled
            InvalidStart, //There are not any roads near the start point of the route
            InvalidEnd, //There are not any roads near the end point of the route
            Failed, //A general route failure occurred.
        }

        #endregion

        #region Structures

        [StructLayout(LayoutKind.Sequential)]
        private struct GPSPositionDataType
        {
            public readonly Int32 lat; //Latitude component of the position in semicircles.
            public readonly Int32 lon; //Longitude component of the position in semicircles.
            public readonly float altMSL; //Altitude above mean sea level component of the position in meters.
            public readonly float altWGS84; //Altitude above WGS84 ellipsoid component of the position in meters.
        };

        [StructLayout(LayoutKind.Sequential)]
        private struct GPSPVTDataType
        {
            public readonly GPSStatusDataType status;
            public readonly GPSPositionDataType position;
            public readonly GPSVelocityDataType velocity;
            public readonly GPSTimeDataType time;
        };

        [StructLayout(LayoutKind.Sequential)]
        private struct GPSSatDataType
        {
            public readonly byte svid; //The space vehicle identifier for the satellite.
            public readonly gpsSat status; //The status bitfield the for satellite (see constants).
            public readonly Int16 snr; //The satellite signal to noise ratio * 100 (dB Hz).
            public readonly float azimuth; //The satellite azimuth (radians).
            public readonly float elevation; //The satellite elevation (radians).
        };

        [StructLayout(LayoutKind.Sequential)]
        private struct GPSStatusDataType
        {
            public readonly GPSModeT8 mode; //GPS Mode
            public readonly GPSFixT8 fix; //GPS Fix
            public readonly Int16 filler2; //Alignment Padding
            public readonly float epe; //The one-sigma estimated position error in meters.
            public readonly float eph; //The one-sigma horizontal only estimated position error in meters.
            public readonly float epv; //The one-sigma vertical only estimated position error in meters.
        };

        [StructLayout(LayoutKind.Sequential)]
        private struct GPSTimeDataType
        {
            public readonly UInt32 seconds; //Seconds since midnight UTC.
            public readonly UInt32 fracSeconds; //To determine the fractional seconds, divide the value in this field by 2^32.
        };

        [StructLayout(LayoutKind.Sequential)]
        private struct GPSVelocityDataType
        {
            public readonly float east; //The East component of the velocity in meters per second.
            public readonly float north; //The North component of the velocity in meters per second.
            public readonly float up; //The upwards component of the velocity in meters per second.
            public readonly float track; //The horizontal vector of the velocity in radians
            public readonly float speed; //The horizontal speed in meters per second.
        };

        [StructLayout(LayoutKind.Sequential)]
        private struct GPSCarrierPhaseOutputPositionDataType
        {
            public readonly double lat; //Latitude (radians)
            public readonly double lon; //Longitude (radians)
            public readonly double tow; //GPS time of week (sec)
            public readonly float alt; //Ellipsoid altitude (meters)
            public readonly float epe; //Estimated position error (meters)
            public readonly float eph; //Estimated position error, horizontal (meters)
            public readonly float epv; //Estimated position error, verticle (meters)
            public readonly float msl; //mean sea level height (meters)
            public readonly float east; //The East component of the velocity in meters per second.
            public readonly float north; //The North component of the velocity in meters per second.
            public readonly float up; //The upwards component of the velocity in meters per second.
            public readonly UInt32 grmn_days; //GARMIN days (day since December 31, 1989)
            public readonly UInt32 fix; //0 = no fix; 1 = no fix; 2 = 2D; 3 = 3D; 4 = 2D differential; 5 = 3D differential; 6 and greater – not defined
            public readonly UInt32 leap_scnds; //UTC leap seonds
        };

        [StructLayout(LayoutKind.Sequential)]
        private struct GPSSatelliteInstRecordDataType
        {
            public readonly double pr; //Pseudorange (meters)
            public readonly UInt32 cycles; //Number of accumulated cycles
            public readonly UInt16 phse; //Carrier phase, 1/2048 cycle
            public readonly byte svid; //Satellite number (0 – 31
            public readonly byte snr_dbhz; //Satellite strength, snr in dB*Hz
            public readonly bool slp_dtct; //cycle slip detected, 0 = no cycle slip detected, non-zero = cycle slip detected 
            public readonly bool valid; //Pseudorange valid flag, 0 = information not valid, non-zero = information valid
        };

        [StructLayout(LayoutKind.Sequential)]
        public struct QueSelectAddressType
        {
            [MarshalAsAttribute(UnmanagedType.LPWStr)]
            public string streetAddress;
            [MarshalAsAttribute(UnmanagedType.LPWStr)]
            public string city;
            [MarshalAsAttribute(UnmanagedType.LPWStr)]
            public string state;
            [MarshalAsAttribute(UnmanagedType.LPWStr)]
            public string country;
            [MarshalAsAttribute(UnmanagedType.LPWStr)]
            public string postalCode;
        };

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct QueAddressType
        {
            [MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst = 51)]
            public string streetAddress;
            [MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst = 51)]
            public string city;
            [MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst = 51)]
            public string state;
            [MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst = 51)]
            public string country;
            [MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst = 12)]
            public string postalCode;
        };

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct QuePositionDataType
        {
            public readonly Int32 lat; //The latitude of the point in semicircles. Semicircles are described in GPS data structure GPSPositionDataType.
            public readonly Int32 lon; //The longitude of the point in semicircles. Semicircles are described in GPS data structure GPSPositionDataType.
            public readonly float altMSL; //The altitude above mean sea level of the point in meters. This field is not used.
        };

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct QuePointType
        {
            //quePointIdLen
            [MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst = 51)]
            public readonly string id;
            public readonly QueSymbolT16 smbl;
            public readonly QuePositionDataType posn;
        };

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct QueRiseSetType
        {
            QueGarminTimeT32 rise; //Rise time in Garmin time formart.
            QueGarminTimeT32 set; //Set time in Garmin time formart.
            byte is_day; //Non-zero if day, zero if night.

        };

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct QueRouteInfoType
        {
            QueRouteStatusT8 routeStatus; //Status of the route.
            float distanceToTurn; //Distance to the next turn in meters.
            float distanceToDest; //Distance to the destination in meters
            QueGarminTimeT32 timeOfTurn; //Estimated time of the next turn
            QueGarminTimeT32 timeOfArrival; //Estimated time of arrival
            [MarshalAsAttribute(UnmanagedType.ByValTStr, SizeConst = 41)]
            string destName; //A pointer to a NULL-terminated string containing the name of the destination
        };

        #endregion

        #region QueAPI

        private delegate void QueNotificationCallback(QueNotificationT8 aNotification);

        //Opens the Que API Library and prepares it for use. 
        // Called by any application or library that wants to use the services that the library provides.
        //QueAPIOpen() must be called before calling any other Que API Library functions, with the exception of QueGetAPIVersion.
        //If the return value is anything other than queErrNone the library was not opened.
        //The application can register GPS data notification by supplying callback function
        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 QueAPIOpen(QueNotificationCallback callback);

        //Closes the QueAPI Library and disposes of the global data memory if required.
        //Called by any application or library that's been using the QueAPI  Library and is now finished with it. 
        //Supply the callback function supplied in the corresponding QueAPIOpen call.
        //This should not be called if GPSOpen failed
        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 QueAPIClose(QueNotificationCallback callback);

        //The API version of the library multiplied by 100.
        //For example, version 1.10 will be returned as 110. 
        //If the version of the library is less than you expect, 
        //it is likely not safe to use, as some functions may not be available
        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern UInt16 QueGetAPIVersion();

        //Can be called without opening the QueAPI Library first.
        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 QueLaunchApp(QueAppT8 app);

        //The value returned by this routine should be used in the dynamic allocation of the array of satellites (GPSSatDataType).
        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern byte GPSGetMaxSatellites();

        //If the return value is not gpsErrNone, the data should be considered invalid.
        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 GPSGetPosition(ref GPSPositionDataType position);

        //Get current position, velocity, and time data.
        //If the return value is not gpsErrNone, the data should be considered invalid.
        //If pvt->status.fix is equal to gpsFixUnusable or gpsFixInvalid, the rest of the data in the structure should be considered invalid.
        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 GPSGetPVT(ref GPSPVTDataType pvt);

        //Get current satellite data.
        //If the return value is not gpsErrNone, the data should be considered invalid.
        //The sat parameter must point to enough memory to hold the maximum number of satellites worth of satellite data.
        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 GPSGetSatellites(ref GPSSatDataType sat);

        //Get current status data.
        //If the return value is not gpsErrNone, the data should be considered invalid.
        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 GPSGetStatus(ref GPSStatusDataType status);

        //Get current time data.
        //If the return value is not gpsErrNone, the data should be considered invalid.
        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 GPSGetTime(ref GPSTimeDataType time);

        //Get current velocity data.
        //If the return value is not gpsErrNone, the data should be considered invalid.
        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 GPSGetVelocity(ref GPSVelocityDataType velocity);

        //Get current PVT data as carrier phase output.
        //If the return value is not gpsErrNone, the data should be considered invalid.
        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 GPSGetCPOPositionData(ref GPSCarrierPhaseOutputPositionDataType position);

        //Get current satellite receiver measurement data from the GPS.
        //If the return value is not gpsErrNone, the data should be considered invalid.
        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 GPSGetSatelliteInstRecordData(ref double rcvr_tow, ref UInt16 rcvr_wn, ref GPSSatelliteInstRecordDataType sats_inst);

        //Set the current GPS mode.
        //If the return value is not gpsErrNone, the mode was not changed.
        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 GPSSetMode(GPSModeT8 mode);

        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 QueCreatePoint(QuePointType pointData, ref QuePointHandle point);

        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 QueCreatePointFromAddress(ref QueSelectAddressType address, ref QuePointHandle point);

        //Returns the serialized data (i.e. series of bytes) that represents the point. This is used for long-term storage of the point. 
        //The point can be re-created by calling QueDeserializePoint().
        //This always returns the size in bytes of the serialized data. If the supplied buffer is not large enough to hold all the serialized data, 
        //no data will be written into the buffer.
        //Typical usage is to call QueSerializePoint() once with pointData set to NULL and pointDataSize set to 0, 
        //then use the returned size to allocate a buffer to hold the serialized data.
        //Then call QueSerializePoint()again with the address and size of the allocated buffer.
        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern UInt32 QueSerializePoint(QuePointHandle point, IntPtr pointData, UInt32 pointDataSize);

        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 QueDeserializePoint(IntPtr pointData, UInt32 pointDataSize, ref QuePointHandle point);

        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 QueGetPointInfo(QuePointHandle point, ref QuePointType pointInfo);

        //Avoid the next metres
        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 QueRouteDetour(float distance);

        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 QueRouteDetour(QueRouteInfoType aRouteInformation);

        //Get current address data in a parsed form
        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 QueGetAddressString(ref QueAddressType aAddress);

        //Get current address data (nearest city).
        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 QueGetAddressString(byte[] aAddress, UInt16 aStringLength);

        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 QueGetDrvRteStatusString(byte[] aStatus, UInt16 aStringLentgh);

        //Get text string for given location.
        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 QueGetStringFromLocation(ref QuePositionDataType aPosn, byte[] aString, UInt16 aStringLength);

        //Get sunrise and sunset times for given location.
        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 QueGetSunRiseSet(ref QuePositionDataType aPosn, ref QueGarminTimeT32 aDate, ref QueRiseSetType aRiseSet);

        //Get moonrise and moonset times for given location.
        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 QueGetMoonRiseSet(ref QuePositionDataType aPosn, ref QueGarminTimeT32 aDate, ref QueRiseSetType aRiseSet);

        //Convert from Garmin time to system time.
        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 QueConvertGarminToSystemTime(QueGarminTimeT32 input, SYSTEMTIME output);

        //Convert from system time to Garmin time.
        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 QueConvertSystemToGarminTime(SYSTEMTIME input, ref QueGarminTimeT32 output);

        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 QueRouteToPoint(QuePointHandle point);

        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 QueRouteIsActive(ref bool active);

        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 QueRouteStop();

        //Creates a route from the current location to a series of points.
        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 QueRouteToPoint(ref QuePointHandle[] points, UInt32 point_count, QueRouteSortT8 sort_type);

        //Displays the QueFind address form to allow the user to select an address from which to create a point.
        //The fields of the address form will be pre-filled with the supplied address data. Not all fields of the input address data need to be supplied; any unused fields should be set to NULL.
        //This call will first attempt to create a point at the location of the specified address exactly like QueCreatePointFromAddress(). 
        //If a single address match is found, it will be returned and the QueFind address form will not be displayed.
        //If a single address match cannot be found, then the QueFind address form is displayed
        //If the user cancels finding an address an invalid point handle will be returned.
        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 QueSelectAddressFromFind(HWND parent, ref QueSelectAddressType address, ref QuePointHandle point);

        //Allows the user to create a point by selecting an item using QueFind.
        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 QueSelectPointFromFind(HWND parent, ref QuePointHandle point);

        //Allows the user to create a point by selecting it from a map.
        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 QueSelectPointFromMap(HWND parent, QuePointHandle orig, ref QuePointHandle point);

        //Switches to the QueMap application centered on the point
        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 QueViewPointOnMap(QuePointHandle point);

        //Displays a modal form containing a map and other details about the point.
        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 QueViewPointDetails(QuePointHandle point);

        //Closes the handle to a point.
        //This must be called for all open point handles before exiting your application. 
        [DllImport("QueAPI.DLL", CallingConvention = CallingConvention.Cdecl)]
        private static extern QueErrT16 QueClosePoint(QuePointHandle point);

        #endregion

        public delegate void GPSEventHandler(object source, QueNotificationT8 evnt);
        public event GPSEventHandler OnGPSEvent;

        private QueNotificationCallback QueNotification;

        private void QueCallback(QueNotificationT8 aNotification)
        {
            if (OnGPSEvent != null)
            {
                OnGPSEvent(this, aNotification);
            }
        }

        protected void Dispose(bool disposing)
        {
            if (disposing)
            {
                QueAPIClose(QueNotification);
                Debug.WriteLine("gps disposed");
            }
        }

        public void Dispose()
        {
            //Do not change this code. 
            //Put cleanup code in
            //Dispose(bool disposing) above.
            Dispose(true);
        }

        ~QueAPI()
        {
            Dispose(false);
        }

        public QueAPI()
        {
            QueNotification = QueCallback;
            if (QueAPIOpen(QueNotification) != QueErrT16.None)
            {
                //throw new UnableToConnect("Cannot connect to Que API");
            }
        }

        //Is the app open?
        public bool isQueOpen()
        {
            bool isActive = false;
            return (QueRouteIsActive(ref isActive) == QueErrT16.None);
        }

        public void openQue(QueAppT8 app)
        {
            QueLaunchApp(app);
        }

        public void navigateToAddress(QueSelectAddressType address)
        {
            QuePointHandle point = new QuePointHandle();
            bool isActive = false;

            if (QueRouteIsActive(ref isActive) == QueErrT16.None)
            {
                if (isActive)
                {
                    QueRouteStop();
                }

                if (QueCreatePointFromAddress(ref address, ref point) == QueErrT16.None)
                {
                    QueRouteToPoint(point);
                    QueClosePoint(point);
                }
            }
        }
    }
}

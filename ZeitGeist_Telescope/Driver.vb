'tabs=4
' --------------------------------------------------------------------------------
' TODO fill in this information for your driver, then remove this line!
'
' ASCOM Telescope driver for ZeitGeist.eq5
'
' Description:	Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam 
'				nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam 
'				erat, sed diam voluptua. At vero eos et accusam et justo duo 
'				dolores et ea rebum. Stet clita kasd gubergren, no sea takimata 
'				sanctus est Lorem ipsum dolor sit amet.
'
' Implements:	ASCOM Telescope interface version: 1.0
' Author:		(XXX) Your N. Here <your@email.here>
'
' Edit Log:
'
' Date			Who	Vers	Description
' -----------	---	-----	-------------------------------------------------------
' dd-mmm-yyyy	XXX	1.0.0	Initial edit, from Telescope template
' ---------------------------------------------------------------------------------
'
'
' Your driver's ID is ASCOM.ZeitGeist.eq5.Telescope
'
' The Guid attribute sets the CLSID for ASCOM.DeviceName.Telescope
' The ClassInterface/None addribute prevents an empty interface called
' _Telescope from being created and used as the [default] interface
'

' This definition is used to select code that's only applicable for one device type
#Const Device = "Telescope"

Imports ASCOM
Imports ASCOM.Astrometry
Imports ASCOM.Astrometry.AstroUtils
Imports ASCOM.DeviceInterface
Imports ASCOM.Utilities

Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Text

<Guid("12622f7f-94be-4485-a98d-cb746dc99faf")> _
<ClassInterface(ClassInterfaceType.None)> _
Public Class Telescope

    ' The Guid attribute sets the CLSID for ASCOM.ZeitGeist.eq5.Telescope
    ' The ClassInterface/None addribute prevents an empty interface called
    ' _ZeitGeist.eq5 from being created and used as the [default] interface

    ' TODO Replace the not implemented exceptions with code to implement the function or
    ' throw the appropriate ASCOM exception.
    '
    Implements ITelescopeV3

    '
    ' Driver ID and descriptive string that shows in the Chooser
    '
    Friend Shared driverID As String = "ASCOM.ZeitGeist.eq5.Telescope"
    Private Shared driverDescription As String = "ZeitGeist.eq5 Telescope"

    Friend Shared comPortProfileName As String = "COM Port" 'Constants used for Profile persistence
    Friend Shared serialSpeedProfileName As String = "Serial Speed" 'Constants used for Profile persistence
    Friend Shared traceStateProfileName As String = "Trace Level"
    Friend Shared followStateProfileName As String = "Follow State"
    Friend Shared comPortDefault As String = "COM1"
    Friend Shared serialSpeedDefault As String = ASCOM.Utilities.SerialSpeed.ps115200
    Friend Shared traceStateDefault As String = "False"
    Friend Shared followStateDefault As String = "False"

    Friend Shared comPort As String ' Variables to hold the currrent device configuration
    Friend Shared serialSpeed As String ' Variables to hold the currrent device configuration
    Friend Shared traceState As Boolean
    Friend Shared followState As Boolean

    Private connectedState As Boolean ' Private variable to hold the connected state
    Private utilities As Util ' Private variable to hold an ASCOM Utilities object
    Private astroUtilities As AstroUtils ' Private variable to hold an AstroUtils object to provide the Range method
    Private TL As TraceLogger ' Private variable to hold the trace logger object (creates a diagnostic log file with information that you specify)

    Private objSerial As ASCOM.Utilities.Serial
    Private isGuidingNow As Boolean = False

    Private property_targetRightAscension As Double
    Private property_targetDeclination As Double

    Private property_rightAscension As Double
    Private property_declination As Double

    Private property_slewing As Boolean = False
    Private property_tracking As Boolean = False

    '
    ' Constructor - Must be public for COM registration!
    '
    Public Sub New()

        ReadProfile() ' Read device configuration from the ASCOM Profile store
        TL = New TraceLogger("", "ZeitGeist.eq5")
        TL.Enabled = traceState
        TL.LogMessage("Telescope", "Starting initialisation")

        connectedState = False ' Initialise connected to false
        utilities = New Util() ' Initialise util object
        astroUtilities = New AstroUtils 'Initialise new astro utiliites object

        'TODO: Implement your additional construction here

        TL.LogMessage("Telescope", "Completed initialisation")
    End Sub

    '
    ' PUBLIC COM INTERFACE ITelescopeV3 IMPLEMENTATION
    '

#Region "Common properties and methods"
    ''' <summary>
    ''' Displays the Setup Dialog form.
    ''' If the user clicks the OK button to dismiss the form, then
    ''' the new settings are saved, otherwise the old values are reloaded.
    ''' THIS IS THE ONLY PLACE WHERE SHOWING USER INTERFACE IS ALLOWED!
    ''' </summary>
    Public Sub SetupDialog() Implements ITelescopeV3.SetupDialog
        ' consider only showing the setup dialog if not connected
        ' or call a different dialog if connected
        If IsConnected Then
            System.Windows.Forms.MessageBox.Show("Already connected, just press OK")
        End If

        Using F As SetupDialogForm = New SetupDialogForm()
            Dim result As System.Windows.Forms.DialogResult = F.ShowDialog()
            If result = DialogResult.OK Then
                WriteProfile() ' Persist device configuration values to the ASCOM Profile store
            End If
        End Using
    End Sub

    Public ReadOnly Property SupportedActions() As ArrayList Implements ITelescopeV3.SupportedActions
        Get
            TL.LogMessage("SupportedActions Get", "Returning empty arraylist")
            Return New ArrayList()
        End Get
    End Property

    Public Function Action(ByVal ActionName As String, ByVal ActionParameters As String) As String Implements ITelescopeV3.Action
        Throw New ActionNotImplementedException("Action " & ActionName & " is not supported by this driver")
    End Function

    Public Sub CommandBlind(ByVal Command As String, Optional ByVal Raw As Boolean = False) Implements ITelescopeV3.CommandBlind
        CheckConnected("CommandBlind")
        ' Call CommandString and return as soon as it finishes
        Me.CommandString(Command, Raw)
        ' or
        Throw New MethodNotImplementedException("CommandBlind")
    End Sub

    Public Function CommandBool(ByVal Command As String, Optional ByVal Raw As Boolean = False) As Boolean _
        Implements ITelescopeV3.CommandBool
        CheckConnected("CommandBool")
        Dim ret As String = CommandString(Command, Raw)
        ' TODO decode the return string and return true or false
        ' or
        Throw New MethodNotImplementedException("CommandBool")
    End Function

    Public Function CommandString(ByVal Command As String, Optional ByVal Raw As Boolean = False) As String _
        Implements ITelescopeV3.CommandString
        CheckConnected("CommandString")
        ' it's a good idea to put all the low level communication with the device here,
        ' then all communication calls this function
        ' you need something to ensure that only one command is in progress at a time
        Throw New MethodNotImplementedException("CommandString")
    End Function

    Public Property Connected() As Boolean Implements ITelescopeV3.Connected
        Get
            TL.LogMessage("Connected Get", IsConnected.ToString())
            Return IsConnected
        End Get
        Set(value As Boolean)
            TL.LogMessage("Connected Set", value.ToString())
            If value = IsConnected Then
                Return
            End If

            If value Then
                connectedState = True
                'comPort = My.Settings.CommPort
                'comPort = "COM10"
                TL.LogMessage("Connected Set", "Connecting to port " + comPort)
                ' TODO connect to the device
                'Dim comPort As String = My.Settings.CommPort
                'Properties.Settings.Default.CommPort
                ''' Change it to Properties.Settings.Default.CommPort
                ''' and in setupdialog add Properties.Settings.Default.CommPort = (string)<the text from the combobox>;
                ''' to where you are saving the settings and doubleclick settings and add CommPort as string to it
                objSerial = New ASCOM.Utilities.Serial
                objSerial.Port = CInt(comPort.Substring(3))
                'objSerial.Speed = ASCOM.Utilities.SerialSpeed.ps9600
                'objSerial.Speed = ASCOM.Utilities.SerialSpeed.ps115200
                'objSerial.Speed = ASCOM.Utilities.SerialSpeed.ps19200
                TL.LogMessage("Connected Set", "Serial speed " + serialSpeed)
                objSerial.Speed = serialSpeed
                objSerial.Connected = True

                Dim startingLoop = True

                While startingLoop
                    Dim serialResponse As String
                    serialResponse = objSerial.ReceiveTerminated("#")
                    serialResponse = serialResponse.Replace("#", "")
                    serialResponse = serialResponse.Replace(vbLf, "")
                    serialResponse = serialResponse.Replace(vbCr, "")
                    serialResponse = serialResponse.Replace(vbCrLf, "")
                    serialResponse = serialResponse.Replace(vbNewLine, "")
                    TL.LogMessage("Connected <- ", serialResponse)

                    If serialResponse = "SYSTEM_STARTED" Then
                        startingLoop = False
                    End If
                End While

                Tracking = followState

            Else
                connectedState = False
                TL.LogMessage("Connected Set", "Disconnecting from port " + comPort)
                ' TODO disconnect from the device
                objSerial.Connected = False
                objSerial.Dispose()
                objSerial = Nothing

            End If
        End Set
    End Property

    Public ReadOnly Property Description As String Implements ITelescopeV3.Description
        Get
            ' this pattern seems to be needed to allow a public property to return a private field
            Dim d As String = driverDescription
            TL.LogMessage("Description Get", d)
            Return d
        End Get
    End Property

    Public ReadOnly Property DriverInfo As String Implements ITelescopeV3.DriverInfo
        Get
            Dim m_version As Version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version
            ' TODO customise this driver description
            Dim s_driverInfo As String = "Information about the driver itself. Version: " + m_version.Major.ToString() + "." + m_version.Minor.ToString()
            TL.LogMessage("DriverInfo Get", s_driverInfo)
            Return s_driverInfo
        End Get
    End Property

    Public ReadOnly Property DriverVersion() As String Implements ITelescopeV3.DriverVersion
        Get
            ' Get our own assembly and report its version number
            TL.LogMessage("DriverVersion Get", Reflection.Assembly.GetExecutingAssembly.GetName.Version.ToString(2))
            Return Reflection.Assembly.GetExecutingAssembly.GetName.Version.ToString(2)
        End Get
    End Property

    Public ReadOnly Property InterfaceVersion() As Short Implements ITelescopeV3.InterfaceVersion
        Get
            TL.LogMessage("InterfaceVersion Get", "3")
            Return 3
        End Get
    End Property

    Public ReadOnly Property Name As String Implements ITelescopeV3.Name
        Get
            Dim s_name As String = "Short driver name - please customise"
            TL.LogMessage("Name Get", s_name)
            Return s_name
        End Get
    End Property

    Public Sub Dispose() Implements ITelescopeV3.Dispose
        ' Clean up the tracelogger and util objects
        TL.Enabled = False
        TL.Dispose()
        TL = Nothing
        utilities.Dispose()
        utilities = Nothing
        astroUtilities.Dispose()
        astroUtilities = Nothing
    End Sub

#End Region

#Region "ITelescope Implementation"
    Public Sub AbortSlew() Implements ITelescopeV3.AbortSlew
        TL.LogMessage("AbortSlew", "Not implemented")
        Throw New ASCOM.MethodNotImplementedException("AbortSlew")
    End Sub

    Public ReadOnly Property AlignmentMode() As AlignmentModes Implements ITelescopeV3.AlignmentMode
        Get
            TL.LogMessage("AlignmentMode Get", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("AlignmentMode", False)
        End Get
    End Property

    Public ReadOnly Property Altitude() As Double Implements ITelescopeV3.Altitude
        Get
            TL.LogMessage("Altitude", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("Altitude", False)
        End Get
    End Property

    Public ReadOnly Property ApertureArea() As Double Implements ITelescopeV3.ApertureArea
        Get
            TL.LogMessage("ApertureArea Get", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("ApertureArea", False)
        End Get
    End Property

    Public ReadOnly Property ApertureDiameter() As Double Implements ITelescopeV3.ApertureDiameter
        Get
            TL.LogMessage("ApertureDiameter Get", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("ApertureDiameter", False)
        End Get
    End Property

    Public ReadOnly Property AtHome() As Boolean Implements ITelescopeV3.AtHome
        Get
            TL.LogMessage("AtHome", "Get - " & False.ToString())
            Return False
        End Get
    End Property

    Public ReadOnly Property AtPark() As Boolean Implements ITelescopeV3.AtPark
        Get
            TL.LogMessage("AtPark", "Get - " & False.ToString())
            Return False
        End Get
    End Property

    Public Function AxisRates(Axis As TelescopeAxes) As IAxisRates Implements ITelescopeV3.AxisRates
        TL.LogMessage("AxisRates", "Get - " & Axis.ToString())
        Return New AxisRates(Axis)
    End Function

    Public ReadOnly Property Azimuth() As Double Implements ITelescopeV3.Azimuth
        Get
            TL.LogMessage("Azimuth Get", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("Azimuth", False)
        End Get
    End Property

    Public ReadOnly Property CanFindHome() As Boolean Implements ITelescopeV3.CanFindHome
        Get
            TL.LogMessage("CanFindHome", "Get - " & False.ToString())
            Return False
        End Get
    End Property

    Public Function CanMoveAxis(Axis As TelescopeAxes) As Boolean Implements ITelescopeV3.CanMoveAxis
        TL.LogMessage("CanMoveAxis", "Get - " & Axis.ToString())
        Select Case Axis
            Case TelescopeAxes.axisPrimary
                Return False
            Case TelescopeAxes.axisSecondary
                Return False
            Case TelescopeAxes.axisTertiary
                Return False
            Case Else
                Throw New InvalidValueException("CanMoveAxis", Axis.ToString(), "0 to 2")
        End Select
    End Function

    Public ReadOnly Property CanPark() As Boolean Implements ITelescopeV3.CanPark
        Get
            TL.LogMessage("CanPark", "Get - " & False.ToString())
            Return False
        End Get
    End Property

    ''' @TODO : This property should be True
    Public ReadOnly Property CanPulseGuide() As Boolean Implements ITelescopeV3.CanPulseGuide
        Get
            'TL.LogMessage("CanPulseGuide", "Get - " & False.ToString())
            'Return False
            TL.LogMessage("CanPulseGuide", "Get - " & True.ToString())
            Return True
        End Get
    End Property

    Public ReadOnly Property CanSetDeclinationRate() As Boolean Implements ITelescopeV3.CanSetDeclinationRate
        Get
            TL.LogMessage("CanSetDeclinationRate", "Get - " & False.ToString())
            Return False
        End Get
    End Property

    Public ReadOnly Property CanSetGuideRates() As Boolean Implements ITelescopeV3.CanSetGuideRates
        Get
            TL.LogMessage("CanSetGuideRates", "Get - " & False.ToString())
            Return False
        End Get
    End Property

    Public ReadOnly Property CanSetPark() As Boolean Implements ITelescopeV3.CanSetPark
        Get
            TL.LogMessage("CanSetPark", "Get - " & False.ToString())
            Return False
        End Get
    End Property

    Public ReadOnly Property CanSetPierSide() As Boolean Implements ITelescopeV3.CanSetPierSide
        Get
            TL.LogMessage("CanSetPierSide", "Get - " & False.ToString())
            Return False
        End Get
    End Property

    Public ReadOnly Property CanSetRightAscensionRate() As Boolean Implements ITelescopeV3.CanSetRightAscensionRate
        Get
            TL.LogMessage("CanSetRightAscensionRate", "Get - " & False.ToString())
            Return False
        End Get
    End Property

    Public ReadOnly Property CanSetTracking() As Boolean Implements ITelescopeV3.CanSetTracking
        Get
            TL.LogMessage("CanSetTracking", "Get - " & True.ToString())
            Return True
        End Get
    End Property

    Public ReadOnly Property CanSlew() As Boolean Implements ITelescopeV3.CanSlew
        Get
            TL.LogMessage("CanSlew", "Get - " & True.ToString())
            Return True
        End Get
    End Property

    Public ReadOnly Property CanSlewAltAz() As Boolean Implements ITelescopeV3.CanSlewAltAz
        Get
            TL.LogMessage("CanSlewAltAz", "Get - " & False.ToString())
            Return False
        End Get
    End Property

    Public ReadOnly Property CanSlewAltAzAsync() As Boolean Implements ITelescopeV3.CanSlewAltAzAsync
        Get
            TL.LogMessage("CanSlewAltAzAsync", "Get - " & False.ToString())
            Return False
        End Get
    End Property

    Public ReadOnly Property CanSlewAsync() As Boolean Implements ITelescopeV3.CanSlewAsync
        Get
            TL.LogMessage("CanSlewAsync", "Get - " & True.ToString())
            Return True
        End Get
    End Property

    'True if this telescope is capable of programmed synching to equatorial coordinates.
    Public ReadOnly Property CanSync() As Boolean Implements ITelescopeV3.CanSync
        Get
            TL.LogMessage("CanSync", "Get - " & True.ToString())
            Return True
        End Get
    End Property

    Public ReadOnly Property CanSyncAltAz() As Boolean Implements ITelescopeV3.CanSyncAltAz
        Get
            TL.LogMessage("CanSyncAltAz", "Get - " & False.ToString())
            Return False
        End Get
    End Property

    Public ReadOnly Property CanUnpark() As Boolean Implements ITelescopeV3.CanUnpark
        Get
            TL.LogMessage("CanUnpark", "Get - " & False.ToString())
            Return False
        End Get
    End Property

    'The declination (degrees) of the telescope's current equatorial coordinates, in the coordinate system given by the EquatorialSystem property. Reading the property will raise an error if the value is unavailable.
    Public ReadOnly Property Declination() As Double Implements ITelescopeV3.Declination
        Get
            refreshDeclination()
            TL.LogMessage("Declination", "Get - " & utilities.DegreesToDMS(property_declination, ":", ":"))
            Return property_declination
        End Get
    End Property

    'The declination tracking rate (arcseconds per SI second, default = 0.0)
    Public Property DeclinationRate() As Double Implements ITelescopeV3.DeclinationRate
        Get
            Dim declination As Double = 0.0
            TL.LogMessage("DeclinationRate", "Get - " & declination.ToString())
            Return declination
        End Get
        Set(value As Double)
            TL.LogMessage("DeclinationRate Set", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("DeclinationRate", True)
        End Set
    End Property

    Public Function DestinationSideOfPier(RightAscension As Double, Declination As Double) As PierSide Implements ITelescopeV3.DestinationSideOfPier
        TL.LogMessage("DestinationSideOfPier Get", "Not implemented")
        Throw New ASCOM.MethodNotImplementedException("DestinationSideOfPier")
    End Function

    'True if the telescope or driver applies atmospheric refraction to coordinates.
    Public Property DoesRefraction() As Boolean Implements ITelescopeV3.DoesRefraction
        Get
            'TL.LogMessage("DoesRefraction Get", "Not implemented")
            'Throw New ASCOM.PropertyNotImplementedException("DoesRefraction", False)
            TL.LogMessage("DoesRefraction Get", "False")
            Return False
        End Get
        Set(value As Boolean)
            TL.LogMessage("DoesRefraction Set", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("DoesRefraction", True)
        End Set
    End Property

    Public ReadOnly Property EquatorialSystem() As EquatorialCoordinateType Implements ITelescopeV3.EquatorialSystem
        Get
            Dim equatorialSystem__1 As EquatorialCoordinateType = EquatorialCoordinateType.equTopocentric
            TL.LogMessage("EquatorialSystem", "Get - " & equatorialSystem__1.ToString())
            Return equatorialSystem__1
        End Get
    End Property

    Public Sub FindHome() Implements ITelescopeV3.FindHome
        TL.LogMessage("FindHome", "Not implemented")
        Throw New ASCOM.MethodNotImplementedException("FindHome")
    End Sub

    Public ReadOnly Property FocalLength() As Double Implements ITelescopeV3.FocalLength
        Get
            TL.LogMessage("FocalLength Get", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("FocalLength", False)
        End Get
    End Property

    Public Property GuideRateDeclination() As Double Implements ITelescopeV3.GuideRateDeclination
        Get
            TL.LogMessage("GuideRateDeclination Get", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("GuideRateDeclination", False)
        End Get
        Set(value As Double)
            TL.LogMessage("GuideRateDeclination Set", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("GuideRateDeclination", True)
        End Set
    End Property

    Public Property GuideRateRightAscension() As Double Implements ITelescopeV3.GuideRateRightAscension
        Get
            TL.LogMessage("GuideRateRightAscension Get", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("GuideRateRightAscension", False)
        End Get
        Set(value As Double)
            TL.LogMessage("GuideRateRightAscension Set", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("GuideRateRightAscension", True)
        End Set
    End Property

    ''' @TODO: this property should be True when PulseGuiding is in progress
    Public ReadOnly Property IsPulseGuiding() As Boolean Implements ITelescopeV3.IsPulseGuiding
        Get
            'TL.LogMessage("IsPulseGuiding Get", "Not implemented")
            'Throw New ASCOM.PropertyNotImplementedException("IsPulseGuiding", False)
            Return isGuidingNow
        End Get
    End Property

    Public Sub MoveAxis(Axis As TelescopeAxes, Rate As Double) Implements ITelescopeV3.MoveAxis
        TL.LogMessage("MoveAxis", "Not implemented")
        Throw New ASCOM.MethodNotImplementedException("MoveAxis")
    End Sub

    Public Sub Park() Implements ITelescopeV3.Park
        TL.LogMessage("Park", "Not implemented")
        Throw New ASCOM.MethodNotImplementedException("Park")
    End Sub

    ''' The main guiding method
    ' @TODO : Implement guiding logic
    ' This method should toggle IsPulseGuiding property
    ' Duration is in milliseconds
    Public Sub PulseGuide(Direction As GuideDirections, Duration As Integer) Implements ITelescopeV3.PulseGuide

        TL.LogMessage("PulseGuide started, Direction ", Duration.ToString)
        TL.LogMessage("PulseGuide started, Duration ", Duration.ToString)

        isGuidingNow = True

        Dim guideDirection As String = ""

        If Direction = GuideDirections.guideNorth Then
            guideDirection = "0"
        End If

        If Direction = GuideDirections.guideSouth Then
            guideDirection = "1"
        End If

        If Direction = GuideDirections.guideEast Then
            guideDirection = "2"
        End If

        If Direction = GuideDirections.guideWest Then
            guideDirection = "3"
        End If

        'Dim commandString = "GUIDE_" + Direction.ToString() + "_" + Duration.ToString
        Dim commandString = "GUIDE_" + guideDirection + "_" + Duration.ToString
        commandString = commandString + "#"

        My.Application.Log.WriteEntry("Going to send to Serial port: " & commandString)

        objSerial.Transmit(commandString)

        Dim serialResponse As String
        serialResponse = objSerial.ReceiveTerminated("#")
        serialResponse = serialResponse.Replace("#", "")
        TL.LogMessage("PulseGuide response ", serialResponse)

        isGuidingNow = False

    End Sub

    'The right ascension (hours) of the telescope's current equatorial coordinates, in the coordinate system given by the EquatorialSystem property
    Public ReadOnly Property RightAscension() As Double Implements ITelescopeV3.RightAscension
        Get
            refreshRightAscension()
            TL.LogMessage("RightAscension", "Get - " & utilities.HoursToHMS(property_rightAscension))
            Return property_rightAscension
        End Get
    End Property

    Public Property RightAscensionRate() As Double Implements ITelescopeV3.RightAscensionRate
        Get
            Dim rightAscensionRate__1 As Double = 0.0
            TL.LogMessage("RightAscensionRate", "Get - " & rightAscensionRate__1.ToString())
            Return rightAscensionRate__1
        End Get
        Set(value As Double)
            TL.LogMessage("RightAscensionRate Set", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("RightAscensionRate", True)
        End Set
    End Property

    Public Sub SetPark() Implements ITelescopeV3.SetPark
        TL.LogMessage("SetPark", "Not implemented")
        Throw New ASCOM.MethodNotImplementedException("SetPark")
    End Sub

    'Indicates the pointing state of the mount.
    '@SEE https://ascom-standards.org/Help/Platform/html/P_ASCOM_DeviceInterface_ITelescopeV3_SideOfPier.htm
    Public Property SideOfPier() As PierSide Implements ITelescopeV3.SideOfPier
        Get
            TL.LogMessage("SideOfPier Get", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("SideOfPier", False)
        End Get
        Set(value As PierSide)
            TL.LogMessage("SideOfPier Set", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("SideOfPier", True)
        End Set
    End Property

    Public ReadOnly Property SiderealTime() As Double Implements ITelescopeV3.SiderealTime
        Get
            ' now using novas 3.1
            Dim lst As Double = 0.0
            Using novas As New ASCOM.Astrometry.NOVAS.NOVAS31
                Dim jd As Double = utilities.DateUTCToJulian(DateTime.UtcNow)
                novas.SiderealTime(jd, 0, novas.DeltaT(jd),
                                   Astrometry.GstType.GreenwichMeanSiderealTime,
                                   Astrometry.Method.EquinoxBased,
                                   Astrometry.Accuracy.Reduced,
                                   lst)
            End Using

            ' Allow for the longitude
            lst += SiteLongitude / 360.0 * 24.0

            ' Reduce to the range 0 to 24 hours
            lst = astroUtilities.ConditionRA(lst)

            TL.LogMessage("SiderealTime", "Get - " & lst.ToString())
            Return lst
        End Get
    End Property

    Public Property SiteElevation() As Double Implements ITelescopeV3.SiteElevation
        Get
            TL.LogMessage("SiteElevation Get", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("SiteElevation", False)
        End Get
        Set(value As Double)
            TL.LogMessage("SiteElevation Set", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("SiteElevation", True)
        End Set
    End Property

    Public Property SiteLatitude() As Double Implements ITelescopeV3.SiteLatitude
        Get
            TL.LogMessage("SiteLatitude Get", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("SiteLatitude", False)
        End Get
        Set(value As Double)
            TL.LogMessage("SiteLatitude Set", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("SiteLatitude", True)
        End Set
    End Property

    Public Property SiteLongitude() As Double Implements ITelescopeV3.SiteLongitude
        Get
            TL.LogMessage("SiteLongitude Get", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("SiteLongitude", False)
        End Get
        Set(value As Double)
            TL.LogMessage("SiteLongitude Set", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("SiteLongitude", True)
        End Set
    End Property

    Public Property SlewSettleTime() As Short Implements ITelescopeV3.SlewSettleTime
        Get
            TL.LogMessage("SlewSettleTime Get", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("SlewSettleTime", False)
        End Get
        Set(value As Short)
            TL.LogMessage("SlewSettleTime Set", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("SlewSettleTime", True)
        End Set
    End Property

    Public Sub SlewToAltAz(Azimuth As Double, Altitude As Double) Implements ITelescopeV3.SlewToAltAz
        TL.LogMessage("SlewToAltAz", "Not implemented")
        Throw New ASCOM.MethodNotImplementedException("SlewToAltAz")
    End Sub

    Public Sub SlewToAltAzAsync(Azimuth As Double, Altitude As Double) Implements ITelescopeV3.SlewToAltAzAsync
        TL.LogMessage("SlewToAltAzAsync", "Not implemented")
        Throw New ASCOM.MethodNotImplementedException("SlewToAltAzAsync")
    End Sub

    Public Sub SlewToCoordinates(RightAscension As Double, Declination As Double) Implements ITelescopeV3.SlewToCoordinates
        TL.LogMessage("SlewToCoordinates", "Not implemented")
        Throw New ASCOM.MethodNotImplementedException("SlewToCoordinates")
    End Sub

    'Move the telescope to the given equatorial coordinates, return immediately after starting the slew.
    Public Sub SlewToCoordinatesAsync(RightAscension As Double, Declination As Double) Implements ITelescopeV3.SlewToCoordinatesAsync
        TL.LogMessage("SlewToCoordinatesAsync", "RightAscension: " & RightAscension.ToString & " ; Declination: " & Declination.ToString)
        TL.LogMessage("- SlewToCoordinatesAsync", "RightAscension: " & utilities.HoursToHMS(RightAscension) & " ; Declination: " & utilities.DegreesToDMS(Declination, "d", "m", "s"))
        TargetRightAscension = RightAscension
        TargetDeclination = Declination
        'property_slewing = True '@TODO This property should be set from the telescope info message

        'Me.property_rightAscension = RightAscension 'Debug only !
        'Me.property_declination = Declination 'Debug only !

        Dim commandString = "GOTO_RA_COORDS_FROM_STRING=" + utilities.HoursToHMS(RightAscension) + ";" + "GOTO_DEC_COORDS_FROM_STRING=" + utilities.DegreesToDMS(Declination, "d", "m", "s") + "#"

        'My.Application.Log.WriteEntry("Going to send to Serial port: " & commandString)
        TL.LogMessage("SlewToCoordinatesAsync -> ", commandString)

        objSerial.Transmit(commandString)
    End Sub

    Public Sub SlewToTarget() Implements ITelescopeV3.SlewToTarget
        TL.LogMessage("SlewToTarget", "Not implemented")
        Throw New ASCOM.MethodNotImplementedException("SlewToTarget")
    End Sub

    Public Sub SlewToTargetAsync() Implements ITelescopeV3.SlewToTargetAsync
        TL.LogMessage("    Public Sub SlewToTargetAsync() Implements ITelescopeV3.SlewToTargetAsync", "Not implemented")
        Throw New ASCOM.MethodNotImplementedException("SlewToTargetAsync")
    End Sub

    ''' Looks like this property is called by autoguider
    Public ReadOnly Property Slewing() As Boolean Implements ITelescopeV3.Slewing
        Get
            refreshSlewing()
            'TL.LogMessage("Slewing Get", "Not implemented")
            'Throw New ASCOM.PropertyNotImplementedException("Slewing", False)
            'Return False
            TL.LogMessage("Slewing Get", property_slewing.ToString)
            Return property_slewing
        End Get
    End Property

    Public Sub SyncToAltAz(Azimuth As Double, Altitude As Double) Implements ITelescopeV3.SyncToAltAz
        TL.LogMessage("SyncToAltAz", "Not implemented")
        Throw New ASCOM.MethodNotImplementedException("SyncToAltAz")
    End Sub

    'Matches the scope's equatorial coordinates to the given equatorial coordinates.
    Public Sub SyncToCoordinates(RightAscension As Double, Declination As Double) Implements ITelescopeV3.SyncToCoordinates
        TL.LogMessage("SyncToCoordinates", "RightAscension: " & RightAscension.ToString & " ; Declination: " & Declination.ToString)
        TL.LogMessage("- SyncToCoordinates", "RightAscension: " & utilities.HoursToHMS(RightAscension) & " ; Declination: " & utilities.DegreesToDMS(Declination, "d", "m", "s"))
        TargetRightAscension = RightAscension
        TargetDeclination = Declination

        Dim commandString = "SET_RA_COORDS_FROM_STRING=" + utilities.HoursToHMS(RightAscension) + ";" + "SET_DEC_COORDS_FROM_STRING=" + utilities.DegreesToDMS(Declination, "d", "m", "s") + "#"

        'My.Application.Log.WriteEntry("Going to send to Serial port: " & commandString)
        TL.LogMessage("SyncToCoordinates -> ", commandString)

        objSerial.Transmit(commandString)
    End Sub

    Public Sub SyncToTarget() Implements ITelescopeV3.SyncToTarget
        TL.LogMessage("SyncToTarget", "Not implemented")
        Throw New ASCOM.MethodNotImplementedException("SyncToTarget")
    End Sub

    'The declination (degrees, positive North) for the target of an equatorial slew or sync operation
    'Setting this property will raise an error if the given value is outside the range -90 to +90 degrees. Reading the property will raise an error if the value has never been set or is otherwise unavailable.
    Public Property TargetDeclination() As Double Implements ITelescopeV3.TargetDeclination
        Get
            Return property_targetDeclination
        End Get
        Set(value As Double)
            property_targetDeclination = value
        End Set
    End Property

    'The right ascension (hours) for the target of an equatorial slew or sync operation
    'Setting this property will raise an error if the given value is outside the range 0 to 24 hours. Reading the property will raise an error if the value has never been set or is otherwise unavailable.
    Public Property TargetRightAscension() As Double Implements ITelescopeV3.TargetRightAscension
        Get
            Return property_targetRightAscension
        End Get
        Set(value As Double)
            property_targetRightAscension = value
        End Set
    End Property

    Public Property Tracking() As Boolean Implements ITelescopeV3.Tracking
        Get
            refreshTracking()
            'TL.LogMessage("Slewing Get", "Not implemented")
            'Throw New ASCOM.PropertyNotImplementedException("Slewing", False)
            'Return False
            TL.LogMessage("Tracking Get", property_tracking.ToString)
            Return property_tracking
        End Get
        Set(value As Boolean)
            'FOLLOW
            Dim commandString = "FOLLOW=" + If(value, "1", "0") + "#"

            TL.LogMessage("Tracking Set -> ", commandString)

            objSerial.ClearBuffers()
            objSerial.Transmit(commandString)
            objSerial.ClearBuffers()

            Dim serialResponse As String
            serialResponse = objSerial.ReceiveTerminated("#")
            serialResponse = serialResponse.Replace("#", "")
            serialResponse = serialResponse.Replace(vbLf, "")
            serialResponse = serialResponse.Replace(vbCr, "")
            serialResponse = serialResponse.Replace(vbCrLf, "")
            serialResponse = serialResponse.Replace(vbNewLine, "")
            TL.LogMessage("Tracking Set <- ", serialResponse)
        End Set
    End Property

    ' @TODO : Implement later
    Public Property TrackingRate() As DriveRates Implements ITelescopeV3.TrackingRate
        Get
            TL.LogMessage("TrackingRate Get", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("TrackingRate", False)
        End Get
        Set(value As DriveRates)
            TL.LogMessage("TrackingRate Set", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("TrackingRate", True)
        End Set
    End Property

    Public ReadOnly Property TrackingRates() As ITrackingRates Implements ITelescopeV3.TrackingRates
        Get
            Dim trackingRates__1 As ITrackingRates = New TrackingRates()
            TL.LogMessage("TrackingRates", "Get - ")
            For Each driveRate As DriveRates In trackingRates__1
                TL.LogMessage("TrackingRates", "Get - " & driveRate.ToString())
            Next
            Return trackingRates__1
        End Get
    End Property

    Public Property UTCDate() As DateTime Implements ITelescopeV3.UTCDate
        Get
            Dim utcDate__1 As DateTime = DateTime.UtcNow
            TL.LogMessage("UTCDate", String.Format("Get - {0}", utcDate__1))
            Return utcDate__1
        End Get
        Set(value As DateTime)
            Throw New ASCOM.PropertyNotImplementedException("UTCDate", True)
        End Set
    End Property

    Public Sub Unpark() Implements ITelescopeV3.Unpark
        TL.LogMessage("Unpark", "Not implemented")
        Throw New ASCOM.MethodNotImplementedException("Unpark")
    End Sub

#End Region

#Region "Private properties and methods"
    ' here are some useful properties and methods that can be used as required
    ' to help with

#Region "ASCOM Registration"

    Private Shared Sub RegUnregASCOM(ByVal bRegister As Boolean)

        Using P As New Profile() With {.DeviceType = "Telescope"}
            If bRegister Then
                P.Register(driverID, driverDescription)
            Else
                P.Unregister(driverID)
            End If
        End Using

    End Sub

    <ComRegisterFunction()> _
    Public Shared Sub RegisterASCOM(ByVal T As Type)

        RegUnregASCOM(True)

    End Sub

    <ComUnregisterFunction()> _
    Public Shared Sub UnregisterASCOM(ByVal T As Type)

        RegUnregASCOM(False)

    End Sub

#End Region

    ''' <summary>
    ''' Returns true if there is a valid connection to the driver hardware
    ''' </summary>
    Private ReadOnly Property IsConnected As Boolean
        Get
            ' TODO check that the driver hardware connection exists and is connected to the hardware
            If connectedState = False Then
                Return False
            End If

            If objSerial Is Nothing Then
                Return False
            End If

            If objSerial.Connected = False Then
                Return False
            End If

            Return True
        End Get
    End Property

    ''' <summary>
    ''' Use this function to throw an exception if we aren't connected to the hardware
    ''' </summary>
    ''' <param name="message"></param>
    Private Sub CheckConnected(ByVal message As String)
        If Not IsConnected Then
            Throw New NotConnectedException(message)
        End If
    End Sub

    ''' <summary>
    ''' Read the device configuration from the ASCOM Profile store
    ''' </summary>
    Friend Sub ReadProfile()
        Using driverProfile As New Profile()
            driverProfile.DeviceType = "Telescope"
            traceState = Convert.ToBoolean(driverProfile.GetValue(driverID, traceStateProfileName, String.Empty, traceStateDefault))
            followState = Convert.ToBoolean(driverProfile.GetValue(driverID, followStateProfileName, String.Empty, followStateDefault))
            comPort = driverProfile.GetValue(driverID, comPortProfileName, String.Empty, comPortDefault)
            serialSpeed = driverProfile.GetValue(driverID, serialSpeedProfileName, String.Empty, serialSpeedDefault)
        End Using
    End Sub

    ''' <summary>
    ''' Write the device configuration to the  ASCOM  Profile store
    ''' </summary>
    Friend Sub WriteProfile()
        Using driverProfile As New Profile()
            driverProfile.DeviceType = "Telescope"
            driverProfile.WriteValue(driverID, traceStateProfileName, traceState.ToString())
            driverProfile.WriteValue(driverID, followStateProfileName, followState.ToString())
            driverProfile.WriteValue(driverID, comPortProfileName, comPort.ToString())
            driverProfile.WriteValue(driverID, serialSpeedProfileName, serialSpeed.ToString())
        End Using

    End Sub

    Private Sub refreshRightAscension()
        'GET_RA_COORDS_AS_STRING
        Dim commandString = "GET_RA_COORDS_AS_STRING" + "#"

        'My.Application.Log.WriteEntry("Going to send to Serial port: " & commandString)
        TL.LogMessage("refreshRightAscension -> ", commandString)

        objSerial.ClearBuffers()
        objSerial.Transmit(commandString)
        objSerial.ClearBuffers()

        Dim serialResponse As String
        serialResponse = objSerial.ReceiveTerminated("#")
        serialResponse = serialResponse.Replace("#", "")
        serialResponse = serialResponse.Replace(vbLf, "")
        serialResponse = serialResponse.Replace(vbCr, "")
        serialResponse = serialResponse.Replace(vbCrLf, "")
        serialResponse = serialResponse.Replace(vbNewLine, "")
        TL.LogMessage("refreshRightAscension <- ", serialResponse)

        Dim responseArray() As String = Split(serialResponse, "=")

        property_rightAscension = utilities.HMSToHours(responseArray(1))
    End Sub

    Private Sub refreshDeclination()
        'GET_DEC_COORDS_AS_STRING
        Dim commandString = "GET_DEC_COORDS_AS_STRING" + "#"

        'My.Application.Log.WriteEntry("Going to send to Serial port: " & commandString)
        TL.LogMessage("refreshDeclination -> ", commandString)

        objSerial.ClearBuffers()
        objSerial.Transmit(commandString)
        objSerial.ClearBuffers()

        Dim serialResponse As String
        serialResponse = objSerial.ReceiveTerminated("#")
        serialResponse = serialResponse.Replace("#", "")
        serialResponse = serialResponse.Replace(vbLf, "")
        serialResponse = serialResponse.Replace(vbCr, "")
        serialResponse = serialResponse.Replace(vbCrLf, "")
        serialResponse = serialResponse.Replace(vbNewLine, "")
        TL.LogMessage("refreshDeclination <- ", serialResponse)

        Dim responseArray() As String = Split(serialResponse, "=")

        property_declination = utilities.DMSToDegrees(responseArray(1))
    End Sub

    Private Sub refreshTracking()
        'GET_FOLLOW
        Dim commandString = "GET_FOLLOW" + "#"

        'My.Application.Log.WriteEntry("Going to send to Serial port: " & commandString)
        TL.LogMessage("refreshTracking -> ", commandString)

        objSerial.ClearBuffers()
        objSerial.Transmit(commandString)
        objSerial.ClearBuffers()

        Dim serialResponse As String
        serialResponse = objSerial.ReceiveTerminated("#")
        serialResponse = serialResponse.Replace("#", "")
        serialResponse = serialResponse.Replace(vbLf, "")
        serialResponse = serialResponse.Replace(vbCr, "")
        serialResponse = serialResponse.Replace(vbCrLf, "")
        serialResponse = serialResponse.Replace(vbNewLine, "")
        TL.LogMessage("refreshTracking <- ", serialResponse)

        Dim responseArray() As String = Split(serialResponse, "=")

        property_tracking = If((responseArray(1) = "1"), True, False)
    End Sub

    Private Sub refreshSlewing()
        'GET_SLEW
        Dim commandString = "GET_SLEW" + "#"

        'My.Application.Log.WriteEntry("Going to send to Serial port: " & commandString)
        TL.LogMessage("refreshSlewing -> ", commandString)

        objSerial.ClearBuffers()
        objSerial.Transmit(commandString)
        objSerial.ClearBuffers()

        Dim serialResponse As String
        serialResponse = objSerial.ReceiveTerminated("#")
        serialResponse = serialResponse.Replace("#", "")
        serialResponse = serialResponse.Replace(vbLf, "")
        serialResponse = serialResponse.Replace(vbCr, "")
        serialResponse = serialResponse.Replace(vbCrLf, "")
        serialResponse = serialResponse.Replace(vbNewLine, "")
        TL.LogMessage("refreshSlewing: ", serialResponse)

        Dim responseArray() As String = Split(serialResponse, "=")

        property_slewing = If((responseArray(1) = "1"), True, False)
    End Sub

#End Region

End Class

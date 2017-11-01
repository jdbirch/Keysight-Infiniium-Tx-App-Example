# -*- coding: utf-8 -*-
"""
Created on Tue Oct 10 14:21:21 2017

@author: Joel Birch - many thanks to other original contributors

This example code is provided as is with no expressed warranty.  It
demonstrates a few of the remote automation capabilities
provided by the N5452A Automated Digital Test Automation Toolkit.

Requirements to successfully run this code:
    1. Infiniium SW installation, on a scope or offline
    2. N6462A DDR4 application
    3. Python 2.7
    4. Pyvisa version 1.6 or greater
    5. N5452A Remote Programming Interface- Download from Keysight web for free
    6. Sample waveforms.  See code for required location, or change code

"""

# First set up the connection to the RPI DLL via the Python for .NET module
import sys
import clr # Import the Common Runtime Library Python module
import visa # Only required to launch the Tx application
import time


# IP address of scope - used to connect the remote API to the platform running
# the application - Must be a network connection
#SCOPE_IP_ADDRESS = "10.178.67.10"
SCOPE_IP_ADDRESS = "localhost"

#Application name must match what is on the menu to launch it
#APP_NAME = "U7238E MIPI D-PHY Test App"
#APP_NAME = "N108xA IEEE802.3 Test App"
APP_NAME = "N6462A/N6462B DDR4 Test App"

# VISA address: Only used if you want to launch the Tx app from your code.
#scope_visa_address = "USB0::0x2A8D::0x904B::MY57160101::0::INSTR"
#typical form for local FlexDCA
scope_visa_address = "TCPIP0::{}::hislip0::INSTR".format(SCOPE_IP_ADDRESS)
#scope_visa_address = "TCPIP0::localhost::hislip0::INSTR"


def GetSWName(scope):
    '''Returns the IDN string from the scope'''
    result = scope.query("*IDN?")
    return result.strip()


def SimpleMessageEventHandler(source, args):
    '''Event callback handler for Simple Message type'''
    print("Received simple event message: '{}'").format(args.Message)
    print("\nMessageID = {}").format(args.ID)
    if (args.ID == '37437505-160C-4cc8-BA06-093C12994C1E'):
        print("Event Handler: processing 'Confirmation Required' message")
        args.Response = DialogResult.OK
    elif (args.ID == '879629E6-78FA-4a87-B247-A9DB4F0D7330'):
        print("Connection Change Event: Replying with Retry\n")
        args.Response = DialogResult.Retry
    else:
        print('Message not handled')


def GenericMessageEventHandler(source, args):
    '''Event callback handler for Generic Message type'''
    print("Received generic event message: {}").format(args.Message)
    print("MessageID = {}").format(args.ID)
    args.Response = DialogResult.OK


if __name__ == "__main__":

    #Path to specific app remote framework
    sys.path.append("C:/ProgramData/Keysight/DigitalTestApps/Remote \
Toolkit/Version 5.30/Tools")

    # Create a reference from CLR to the Compliance App DLL
    clr.AddReference("Keysight.DigitalTestApps.Framework.Remote")
    clr.AddReference("System.Windows")
    clr.AddReference("System.Runtime")
    clr.AddReference("System.Windows.Forms")
    clr.AddReference("System.Runtime.Remoting")

    # Import the entire compliance app namespace from the Compliance App DLL
    from Keysight.DigitalTestApps.Framework.Remote import *
    # Import the entire .NET remoting namespace
    from System.Runtime.Remoting import *
    from System.Windows.Forms import *

    # The following code will establish a VISA connection and launch the Tx App
    if (float(visa.__version__) < 1.6):
        print ("Error: Expecting pyVISA version 1.6 or greater")
        sys.exit(1)
    rm = visa.ResourceManager('C:/Program Files (x86)/IVI Foundation/VISA/\
WinNT/agvisa/agbin/visa32.dll')
    try:
    # scope is only used to launch the application
        scope = rm.open_resource(scope_visa_address)
        print("VISA connection established to {}\n").format(GetSWName(scope))
    except:
        raise Exception("Connection to instrument at {} failed\n".format\
                        (scope_visa_address))

    # Optional: Launch the compliance app using the pyvisa instrument
    #connection, or launch the application manually before running this code
    # See Remote App Programming Getting Started Guide
    print("Starting Test Application.  Please wait...")
    launchCmd = ":SYSTem:LAUNch \'{}\'".format(APP_NAME)
    scope.write(launchCmd)

    # Connect the Remote Interface to the Digital Test App
    # This commented out non-custom method works well when the app is launched
    # prior to execution.  Some apps take a long time to launch,
    # so in that case, GetRemoteAteCustom() allows a custom timeout
    #remote_obj = RemoteAteUtilities.GetRemoteAte(SCOPE_IP_ADDRESS)
    remoteAteOptions = GetRemoteAteOptions()
    remoteAteOptions.IpAddress = SCOPE_IP_ADDRESS
    remoteAteOptions.TimeoutInSeconds = 120
    remoteAteOptions.UseCustomPort = False
    remoteObj = RemoteAteUtilities.GetRemoteAteCustom(remoteAteOptions)
    remoteApp = IRemoteAte(remoteObj)

    #Verify Connection
    print("Connected to Application: {}\n").format(remoteApp.ApplicationName)

    # Set Application specific configuration properties - See the specific Tx
    # Application remote interface documentation, found on the app Help menu,
    # or sometimes on the product webpage, Document Library tab

	# Setup Tab parameters
    remoteApp.SetConfig("pcboOverallDeviceID", "My Device")
    remoteApp.SetConfig("pcboOverallDeviceDescription", "User Me")
    remoteApp.SetConfig("DeviceType", "DDR4-2400")
    #remoteApp.SetConfig("BurstTrigMethod", "DQS-DQ Phase Difference")
    remoteApp.SetConfig("BurstTrigMethod", "Rd or Wrt ONLY")

    # Setup Offline files
    offlinePath = \
    "C:/Users/Public/Documents/Infiniium/Waveforms/DDR4_2400_SampleWfms"
    clkFile = "{}/{}".format(offlinePath, "clk.wfm")
    dqFile = "{}/{}".format(offlinePath, "dq.wfm")
    dqsFile = "{}/{}".format(offlinePath, "dqs.wfm")

    remoteApp.SetConfig("OfflineClockFilePath", clkFile)
    remoteApp.SetConfig("OfflineDQFilePath", dqFile)
    remoteApp.SetConfig("OfflineDQSFilePath", dqsFile)
    remoteApp.SetConfig("OfflineDataMode", "1")

    # Config Tab parameters
    # Let's check a current value of a config parameter
    configParam = "MaxBurstLenLimit"
    print("Current value of {}:{}\n").format(configParam, \
         remoteApp.GetConfig(configParam))

    # testInfos: All the available tests with the current setup/configuration
    # Available tests changes based on configuration, at least in offline mode
    testsInfos = remoteApp.GetCurrentOptions("TestsInfo")
    print("{} tests available in this configuration\n".format(len(testsInfos)))

    # Build a dictionary of tests with test IDs as keys.  Used a little later
    testID = {}
    for test in testsInfos:
        testID[test.ID] = [test.Name, test.Description.replace("\r", "\n"), \
              test.Reference]
        '''print( u"ID:{} || Name:{}\n  Description: {}\n   Reference: {}\n"\
              .format(test.ID, test.Name, \
                      test.Description.replace("\r", "\n"), test.Reference))
        '''

    #Before running: Establish callback path to handle some pop-up dialogs
    configFilePath = "c:/ProgramData/Keysight/DigitalTestApps/Remote \
Toolkit/Version 5.30/Tools/Keysight.DigitalTestApps.Framework.Remote.config"
    try:
        RemotingConfiguration.Configure(configFilePath, False)
    except RemotingException:
        print("Remoting Configuration for callbacks previously set up")

    eventSink = RemoteAteUtilities.CreateAteEventSink(remoteApp, None, \
                                                      SCOPE_IP_ADDRESS)

    #Event Handling setup.  See Event handler functions too
    #Subscribe to message events
    eventSink.RedirectMessagesToClient = True
    eventSink.SimpleMessageEvent += SimpleMessageEventHandler
    eventSink.GenericMessageEvent += GenericMessageEventHandler


    #InfiniiSim setup example
    # Check for license first
    scopeOptions = scope.query("*OPT?")
    #Doesn't behave well offline, so here just for example
    if scopeOptions.find("xyz") >= 0: # DEA is real option, change to enable
        # Create an InfiniiSim Options object
        isimOpts = InfiniiSimOptions()
        # Set options as desired - See InfiniiSimOptions Properties
        # See InfiniiSimOptions.InfiniiSimState Enumeration
        isimOpts.State = isimOpts.InfiniiSimState.TwoPort
        isimOpts.Bandwidth = 12e9
        isimOpts.TransferFunction = \
        "C:/Users/Public/Documents/Infiniium/Filters/DoNothing.tf2"
        # See InfiniiSimOptions.InfiniiSimPortExtraction
        isimOpts.PortExtraction = isimOpts.InfiniiSimPortExtraction.Ports12
        # Apply options (settings) to channel 1
        remoteApp.SetInfiniiSimSettings(1, isimOpts)
    else:
        print("InfiniiSim Advanced not licensed. Ignoring InfiniiSim setup.")


    # Because we're handling the connection diagram with a callback, the
    # following line is not needed
    #remoteApp.ConnectionPromptAction = 1 #Auto respond - No connection diagram
    #testsToRun = [500, 30104, 30105]
    testsToRun = [500, 30104]
    remoteApp.SelectedTests = testsToRun  # Select the tests to run
    print("\n{} tests selected in this configuration.".format(len(testsToRun)))
    for test in testsToRun:
        print( u"ID:{} || Name:{}\n  Description: {}\n   Reference: {}\n"\
              .format(test, testID[test][0], testID[test][1], testID[test][2]))

    remoteApp.Run() # Run the selected tests

    # Set up the project save options & save it to disk
    saveOptions = SaveProjectOptions()
    saveOptions.BaseDirectory = "c:/temp"
    saveOptions.Name = "Demo"
    saveOptions.OverwriteExisting = True
    projectFullPath = remoteApp.SaveProjectCustom(saveOptions)
    print("Project saved to {}/{}".format(saveOptions.BaseDirectory, \
          saveOptions.Name))

    # Set up the project results options and then get the results
    resultOptions = ResultOptions()
    resultOptions.TestIds = testsToRun
    resultOptions.IncludeCsvData = True
    customResults = remoteApp.GetResultsCustom(resultOptions)
    results = customResults.CsvResults
    csvFile = "c:/temp/results.csv"
    f = open(csvFile, 'w')
    f.write(results)
    f.close()
    print("csv results saved to {}".format(csvFile))
    print("All done.  Killing app")
    remoteApp.Exit(True,True) # Exit the application

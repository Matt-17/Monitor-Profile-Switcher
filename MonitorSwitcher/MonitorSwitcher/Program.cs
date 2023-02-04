/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.IO;

namespace MonitorSwitcherGUI
{
    public class MonitorSwitcher
    {
        private static Boolean debug;
        private static Boolean noIDMatch;

        public static void DebugOutput(String text)
        {
            if (debug)
            {
                Console.WriteLine(text);
            }
        }

        public static Boolean LoadDisplaySettings(String fileName)
        {
            DebugOutput("Loading display settings from file: " + fileName);
            if (!File.Exists(fileName))
            {
                Console.WriteLine("Failed to load display settings because file does not exist: " + fileName);

                return false;
            }

            // Objects for DeSerialization of pathInfo and modeInfo classes
            DebugOutput("Initializing objects for Serialization");
            System.Xml.Serialization.XmlSerializer readerAdditionalInfo = new System.Xml.Serialization.XmlSerializer(typeof(CCDWrapper.MonitorAdditionalInfo));
            System.Xml.Serialization.XmlSerializer readerPath = new System.Xml.Serialization.XmlSerializer(typeof(CCDWrapper.DisplayConfigPathInfo));
            System.Xml.Serialization.XmlSerializer readerModeTarget = new System.Xml.Serialization.XmlSerializer(typeof(CCDWrapper.DisplayConfigTargetMode));
            System.Xml.Serialization.XmlSerializer readerModeSource = new System.Xml.Serialization.XmlSerializer(typeof(CCDWrapper.DisplayConfigSourceMode));
            System.Xml.Serialization.XmlSerializer readerModeInfoType = new System.Xml.Serialization.XmlSerializer(typeof(CCDWrapper.DisplayConfigModeInfoType));
            System.Xml.Serialization.XmlSerializer readerModeAdapterID = new System.Xml.Serialization.XmlSerializer(typeof(CCDWrapper.LUID));

            // Lists for storing the results
            List<CCDWrapper.DisplayConfigPathInfo> pathInfoList = new List<CCDWrapper.DisplayConfigPathInfo>();
            List<CCDWrapper.DisplayConfigModeInfo> modeInfoList = new List<CCDWrapper.DisplayConfigModeInfo>();
            List<CCDWrapper.MonitorAdditionalInfo> additionalInfoList = new List<CCDWrapper.MonitorAdditionalInfo>();

            // Loading the xml file
            DebugOutput("Parsing XML file");
            XmlReader xml = XmlReader.Create(fileName);
            xml.Read();
            while (true)
            {
                DebugOutput("\tXML Element: " + xml.Name);
                if ((xml.Name.CompareTo("DisplayConfigPathInfo") == 0) && (xml.IsStartElement()))
                {
                    CCDWrapper.DisplayConfigPathInfo pathInfo = (CCDWrapper.DisplayConfigPathInfo)readerPath.Deserialize(xml);
                    pathInfoList.Add(pathInfo);
                    continue;
                }
                else if ((xml.Name.CompareTo("modeInfo") == 0) && (xml.IsStartElement()))
                {
                    DebugOutput("\t\tReading modeInfo");
                    CCDWrapper.DisplayConfigModeInfo modeInfo = new CCDWrapper.DisplayConfigModeInfo();
                    xml.Read();
                    xml.Read();
                    modeInfo.id = Convert.ToUInt32(xml.Value);                    
                    xml.Read();
                    xml.Read();
                    modeInfo.adapterId = (CCDWrapper.LUID)readerModeAdapterID.Deserialize(xml);
                    modeInfo.infoType = (CCDWrapper.DisplayConfigModeInfoType)readerModeInfoType.Deserialize(xml);
                    if (modeInfo.infoType == CCDWrapper.DisplayConfigModeInfoType.Target)
                    {
                        modeInfo.targetMode = (CCDWrapper.DisplayConfigTargetMode)readerModeTarget.Deserialize(xml);
                    }
                    else
                    {
                        modeInfo.sourceMode = (CCDWrapper.DisplayConfigSourceMode)readerModeSource.Deserialize(xml);
                    }
                    DebugOutput("\t\t\tmodeInfo.id = " + modeInfo.id);
                    DebugOutput("\t\t\tmodeInfo.adapterId (High Part) = " + modeInfo.adapterId.HighPart);
                    DebugOutput("\t\t\tmodeInfo.adapterId (Low Part) = " + modeInfo.adapterId.LowPart);
                    DebugOutput("\t\t\tmodeInfo.infoType = " + modeInfo.infoType);

                    modeInfoList.Add(modeInfo);
                    continue;
                } else if ((xml.Name.CompareTo("MonitorAdditionalInfo") == 0) && (xml.IsStartElement()))
                {
                    DebugOutput("\t\tReading additional informations");
                    CCDWrapper.MonitorAdditionalInfo additionalInfo = (CCDWrapper.MonitorAdditionalInfo)readerAdditionalInfo.Deserialize(xml);
                    additionalInfoList.Add(additionalInfo);
                    continue;
                }

                    if (!xml.Read())
                {
                    break;
                }
            }
            xml.Close();
            DebugOutput("Parsing of XML file successful");

            // Convert C# lists to simply array
            DebugOutput("Converting to simple arrays for API compatibility");
            var pathInfoArray = new CCDWrapper.DisplayConfigPathInfo[pathInfoList.Count];
            for (int iPathInfo = 0; iPathInfo < pathInfoList.Count; iPathInfo++)
            {
                pathInfoArray[iPathInfo] = pathInfoList[iPathInfo];
            }

            var modeInfoArray = new CCDWrapper.DisplayConfigModeInfo[modeInfoList.Count];
            for (int iModeInfo = 0; iModeInfo < modeInfoList.Count; iModeInfo++)
            {
                modeInfoArray[iModeInfo] = modeInfoList[iModeInfo];
            }

            // Get current display settings
            DebugOutput("Getting current display settings");
            CCDWrapper.DisplayConfigPathInfo[] pathInfoArrayCurrent = new CCDWrapper.DisplayConfigPathInfo[0];
            CCDWrapper.DisplayConfigModeInfo[] modeInfoArrayCurrent = new CCDWrapper.DisplayConfigModeInfo[0];
            CCDWrapper.MonitorAdditionalInfo[] additionalInfoCurrent = new CCDWrapper.MonitorAdditionalInfo[0];

            Boolean statusCurrent = GetDisplaySettings(ref pathInfoArrayCurrent, ref modeInfoArrayCurrent, ref additionalInfoCurrent, false);
            if (statusCurrent)
            {
                if (!noIDMatch)
                {
                    // For some reason the adapterID parameter changes upon system restart, all other parameters however, especially the ID remain constant.
                    // We check the loaded settings against the current settings replacing the adapaterID with the other parameters
                    DebugOutput("Matching of adapter IDs for pathInfo");
                    for (int iPathInfo = 0; iPathInfo < pathInfoArray.Length; iPathInfo++)
                    {
                        for (int iPathInfoCurrent = 0; iPathInfoCurrent < pathInfoArrayCurrent.Length; iPathInfoCurrent++)
                        {
                            DebugOutput("\t---");
                            DebugOutput("\tIndex XML = " + iPathInfo);
                            DebugOutput("\tIndex Current = " + iPathInfoCurrent);
                            DebugOutput("\tsourceInfo.id XML = " + pathInfoArray[iPathInfo].sourceInfo.id);
                            DebugOutput("\tsourceInfo.id Current = " + pathInfoArrayCurrent[iPathInfoCurrent].sourceInfo.id);
                            DebugOutput("\ttargetInfo.id XML = " + pathInfoArray[iPathInfo].targetInfo.id);
                            DebugOutput("\ttargetInfo.id Current = " + pathInfoArrayCurrent[iPathInfoCurrent].targetInfo.id);
                            DebugOutput("\tsourceInfo.adapterId XML = " + pathInfoArray[iPathInfo].sourceInfo.adapterId.LowPart);
                            DebugOutput("\tsourceInfo.adapterId Current = " + pathInfoArrayCurrent[iPathInfoCurrent].sourceInfo.adapterId.LowPart);
                            DebugOutput("\ttargetInfo.adapterId XML = " + pathInfoArray[iPathInfo].targetInfo.adapterId.LowPart);
                            DebugOutput("\ttargetInfo.adapterId Current = " + pathInfoArrayCurrent[iPathInfoCurrent].targetInfo.adapterId.LowPart);
                            if ((pathInfoArray[iPathInfo].sourceInfo.id == pathInfoArrayCurrent[iPathInfoCurrent].sourceInfo.id) &&
                                (pathInfoArray[iPathInfo].targetInfo.id == pathInfoArrayCurrent[iPathInfoCurrent].targetInfo.id))
                            {
                                DebugOutput("\t!!! Both IDs are a match, assigning current adapter ID !!!");
                                pathInfoArray[iPathInfo].sourceInfo.adapterId.LowPart = pathInfoArrayCurrent[iPathInfoCurrent].sourceInfo.adapterId.LowPart;
                                pathInfoArray[iPathInfo].targetInfo.adapterId.LowPart = pathInfoArrayCurrent[iPathInfoCurrent].targetInfo.adapterId.LowPart;
                                break;
                            }
                            DebugOutput("\t---");
                        }
                    }

                    // Same again for modeInfo, however we get the required adapterId information from the pathInfoArray
                    DebugOutput("Matching of adapter IDs for modeInfo");
                    for (int iModeInfo = 0; iModeInfo < modeInfoArray.Length; iModeInfo++)
                    {
                        for (int iPathInfo = 0; iPathInfo < pathInfoArray.Length; iPathInfo++)
                        {
                            DebugOutput("\t---");
                            DebugOutput("\tIndex Mode = " + iModeInfo);
                            DebugOutput("\tIndex Path = " + iPathInfo);
                            DebugOutput("\tmodeInfo.id = " + modeInfoArray[iModeInfo].id);
                            DebugOutput("\tpathInfo.id = " + pathInfoArray[iPathInfo].targetInfo.id);
                            DebugOutput("\tmodeInfo.infoType = " + modeInfoArray[iModeInfo].infoType);
                            if ((modeInfoArray[iModeInfo].id == pathInfoArray[iPathInfo].targetInfo.id) &&
                                (modeInfoArray[iModeInfo].infoType == CCDWrapper.DisplayConfigModeInfoType.Target))
                            {
                                DebugOutput("\t\tTarget adapter id found, checking for source modeInfo and adpaterID");
                                // We found target adapter id, now lets look for the source modeInfo and adapterID
                                for (int iModeInfoSource = 0; iModeInfoSource < modeInfoArray.Length; iModeInfoSource++)
                                {
                                    DebugOutput("\t\t---");
                                    DebugOutput("\t\tIndex = " + iModeInfoSource);
                                    DebugOutput("\t\tmodeInfo.id Source = " + modeInfoArray[iModeInfoSource].id);
                                    DebugOutput("\t\tpathInfo.sourceInfo.id = " + pathInfoArray[iPathInfo].sourceInfo.id);
                                    DebugOutput("\t\tmodeInfo.adapterId = " + modeInfoArray[iModeInfo].adapterId.LowPart);
                                    DebugOutput("\t\tmodeInfo.adapterId Source = " + modeInfoArray[iModeInfoSource].adapterId.LowPart);
                                    DebugOutput("\t\tmodeInfo.infoType Source = " + modeInfoArray[iModeInfoSource].infoType);
                                    if ((modeInfoArray[iModeInfoSource].id == pathInfoArray[iPathInfo].sourceInfo.id) &&
                                        (modeInfoArray[iModeInfoSource].adapterId.LowPart == modeInfoArray[iModeInfo].adapterId.LowPart) &&
                                        (modeInfoArray[iModeInfoSource].infoType == CCDWrapper.DisplayConfigModeInfoType.Source))
                                    {
                                        DebugOutput("\t\t!!! IDs are a match, taking adpater id from pathInfo !!!");
                                        modeInfoArray[iModeInfoSource].adapterId.LowPart = pathInfoArray[iPathInfo].sourceInfo.adapterId.LowPart;
                                        break;
                                    }
                                    DebugOutput("\t\t---");
                                }
                                modeInfoArray[iModeInfo].adapterId.LowPart = pathInfoArray[iPathInfo].targetInfo.adapterId.LowPart;
                                break;
                            }
                            DebugOutput("\t---");
                        }
                    }
                    DebugOutput("Done matching of adapter IDs");
                }

                // Set loaded display settings
                DebugOutput("Setting up final display settings to load");
                if (debug)
                {
                    // debug output complete display settings
                    Console.WriteLine("\nDisplay settings to be loaded: ");
                    Console.WriteLine(PrintDisplaySettings(pathInfoArray, modeInfoArray));
                }
                uint numPathArrayElements = (uint)pathInfoArray.Length;
                uint numModeInfoArrayElements = (uint)modeInfoArray.Length;
                long status = CCDWrapper.SetDisplayConfig(numPathArrayElements, pathInfoArray, numModeInfoArrayElements, modeInfoArray,
                                                          CCDWrapper.SdcFlags.Apply | CCDWrapper.SdcFlags.UseSuppliedDisplayConfig | CCDWrapper.SdcFlags.SaveToDatabase | CCDWrapper.SdcFlags.NoOptimization | CCDWrapper.SdcFlags.AllowChanges);
                if (status != 0)
                {
                    Console.WriteLine("Failed to set display settings using default method, ERROR: " + status.ToString());                    

                    if ((additionalInfoCurrent.Length > 0) && (additionalInfoList.Count > 0)) // only if present, e.g. new profile
                    {
                        Console.WriteLine("Trying alternative method");
                        // Restore original settings and adapter IDs
                        DebugOutput("Converting again to simple arrays for API compatibility");
                        for (int iPathInfo = 0; iPathInfo < pathInfoList.Count; iPathInfo++)
                        {
                            pathInfoArray[iPathInfo] = pathInfoList[iPathInfo];
                        }

                        for (int iModeInfo = 0; iModeInfo < modeInfoList.Count; iModeInfo++)
                        {
                            modeInfoArray[iModeInfo] = modeInfoList[iModeInfo];
                        }

                        DebugOutput("Alternative matching mode");
                        // For each modeInfo iterate over the current additional informations, i.e. monitor names and paths, and find the one matching in the current setup
                        for (int iModeInfo = 0; iModeInfo < modeInfoArray.Length; iModeInfo++)
                        {
                            for (int iAdditionalInfoCurrent = 0; iAdditionalInfoCurrent < additionalInfoCurrent.Length; iAdditionalInfoCurrent++)
                            {
                                if ((additionalInfoCurrent[iAdditionalInfoCurrent].monitorFriendlyDevice != null) && (additionalInfoList[iModeInfo].monitorFriendlyDevice != null))
                                {
                                    if (additionalInfoCurrent[iAdditionalInfoCurrent].monitorFriendlyDevice.Equals(additionalInfoList[iModeInfo].monitorFriendlyDevice))
                                    {
                                        CCDWrapper.LUID originalID = modeInfoArray[iModeInfo].adapterId;                                        
                                        // now also find all other matching pathInfo modeInfos with that ID and change it
                                        for (int iPathInfo = 0; iPathInfo < pathInfoArray.Length; iPathInfo++)
                                        {
                                            if ((pathInfoArray[iPathInfo].targetInfo.adapterId.LowPart == originalID.LowPart) &&
                                               (pathInfoArray[iPathInfo].targetInfo.adapterId.HighPart == originalID.HighPart))
                                            {
                                                pathInfoArray[iPathInfo].targetInfo.adapterId = modeInfoArrayCurrent[iAdditionalInfoCurrent].adapterId;
                                                pathInfoArray[iPathInfo].sourceInfo.adapterId = modeInfoArrayCurrent[iAdditionalInfoCurrent].adapterId;
                                                pathInfoArray[iPathInfo].targetInfo.id = modeInfoArrayCurrent[iAdditionalInfoCurrent].id;
                                            }
                                        }
                                        for (int iModeInfoFix = 0; iModeInfoFix < modeInfoArray.Length; iModeInfoFix++)
                                        {
                                            if ((modeInfoArray[iModeInfoFix].adapterId.LowPart == originalID.LowPart) &&
                                                (modeInfoArray[iModeInfoFix].adapterId.HighPart == originalID.HighPart))
                                            {
                                                modeInfoArray[iModeInfoFix].adapterId = modeInfoArrayCurrent[iAdditionalInfoCurrent].adapterId;                                                
                                            }
                                        }
                                        modeInfoArray[iModeInfo].adapterId = modeInfoArrayCurrent[iAdditionalInfoCurrent].adapterId;
                                        modeInfoArray[iModeInfo].id = modeInfoArrayCurrent[iAdditionalInfoCurrent].id;

                                        break;
                                    }
                                }
                            }
                        }

                        if (debug)
                        {
                            // debug output complete display settings
                            Console.WriteLine("\nDisplay settings to be loaded: ");
                            Console.WriteLine(PrintDisplaySettings(pathInfoArray, modeInfoArray));
                        }
                        status = CCDWrapper.SetDisplayConfig(numPathArrayElements, pathInfoArray, numModeInfoArrayElements, modeInfoArray,
                                                             CCDWrapper.SdcFlags.Apply | CCDWrapper.SdcFlags.UseSuppliedDisplayConfig | CCDWrapper.SdcFlags.NoOptimization | CCDWrapper.SdcFlags.SaveToDatabase | CCDWrapper.SdcFlags.AllowChanges);
                        if (status != 0)
                        {
                            Console.WriteLine("Failed to set display settings using alternative method, ERROR: " + status.ToString());
                            return false;
                        }
                        return true;
                    }
                    return false;
                }
                
                return true;
            }

            DebugOutput("Failed to get current display settings");
            return false;
        }

        public static Boolean GetDisplaySettings(ref CCDWrapper.DisplayConfigPathInfo[] pathInfoArray, ref CCDWrapper.DisplayConfigModeInfo[] modeInfoArray, ref CCDWrapper.MonitorAdditionalInfo[] additionalInfo, Boolean ActiveOnly)
        {
            uint numPathArrayElements;
            uint numModeInfoArrayElements;

            // query active paths from the current computer.
            DebugOutput("Getting display settings");
            CCDWrapper.QueryDisplayFlags queryFlags = CCDWrapper.QueryDisplayFlags.AllPaths;
            if (ActiveOnly)
            {
                queryFlags = CCDWrapper.QueryDisplayFlags.OnlyActivePaths;
            }

            DebugOutput("Getting buffer size");
            var status = CCDWrapper.GetDisplayConfigBufferSizes(queryFlags, out numPathArrayElements, out numModeInfoArrayElements);
            if (status == 0)
            {
                pathInfoArray = new CCDWrapper.DisplayConfigPathInfo[numPathArrayElements];
                modeInfoArray = new CCDWrapper.DisplayConfigModeInfo[numModeInfoArrayElements];
                additionalInfo = new CCDWrapper.MonitorAdditionalInfo[numModeInfoArrayElements];

                DebugOutput("Querying display config");
                status = CCDWrapper.QueryDisplayConfig(queryFlags,
                                                       ref numPathArrayElements, pathInfoArray, ref numModeInfoArrayElements,
                                                       modeInfoArray, IntPtr.Zero);                    

                if (status == 0)
                {
                    // cleanup of modeInfo bad elements 
                    int validCount = 0;
                    foreach (CCDWrapper.DisplayConfigModeInfo modeInfo in modeInfoArray)
                    {
                        if (modeInfo.infoType != CCDWrapper.DisplayConfigModeInfoType.Zero)
                        {   // count number of valid mode Infos
                            validCount++;
                        }
                    }
                    if (validCount > 0)
                    {   // only cleanup if there is at least one valid element found
                        CCDWrapper.DisplayConfigModeInfo[] tempInfoArray = new CCDWrapper.DisplayConfigModeInfo[modeInfoArray.Count()];
                        modeInfoArray.CopyTo(tempInfoArray, 0);
                        modeInfoArray = new CCDWrapper.DisplayConfigModeInfo[validCount];
                        int index = 0;
                        foreach (CCDWrapper.DisplayConfigModeInfo modeInfo in tempInfoArray)
                        {
                            if (modeInfo.infoType != CCDWrapper.DisplayConfigModeInfoType.Zero)
                            {
                                modeInfoArray[index] = modeInfo;
                                index++;
                            }
                        }
                    }

                    // cleanup of currently not available pathInfo elements
                    validCount = 0;
                    foreach (CCDWrapper.DisplayConfigPathInfo pathInfo in pathInfoArray)
                    {
                        if (pathInfo.targetInfo.targetAvailable)
                        {
                            validCount++;
                        }
                    }
                    if (validCount > 0)
                    {   // only cleanup if there is at least one valid element found
                        CCDWrapper.DisplayConfigPathInfo[] tempInfoArray = new CCDWrapper.DisplayConfigPathInfo[pathInfoArray.Count()];
                        pathInfoArray.CopyTo(tempInfoArray, 0);
                        pathInfoArray = new CCDWrapper.DisplayConfigPathInfo[validCount];
                        int index = 0;
                        foreach (CCDWrapper.DisplayConfigPathInfo pathInfo in tempInfoArray)
                        {
                            if (pathInfo.targetInfo.targetAvailable)
                            {
                                pathInfoArray[index] = pathInfo;
                                index++;
                            }
                        }
                    }

                    // get the display names for all modes
                    for (var iMode = 0; iMode < modeInfoArray.Count(); iMode++)
                    {
                        if (modeInfoArray[iMode].infoType == CCDWrapper.DisplayConfigModeInfoType.Target)
                        {
                            try
                            {
                                additionalInfo[iMode] = CCDWrapper.GetMonitorAdditionalInfo(modeInfoArray[iMode].adapterId, modeInfoArray[iMode].id);
                            }
                            catch (Exception e)
                            {
                                additionalInfo[iMode].valid = false;
                            }
                        }
                    }
                    return true;
                } else
                {
                    DebugOutput("Querying display config failed");
                }
            } else
            {
                DebugOutput("Getting Buffer Size Failed");
            }

            return false;
        }

        public static String PrintDisplaySettings(CCDWrapper.DisplayConfigPathInfo[] pathInfoArray, CCDWrapper.DisplayConfigModeInfo[] modeInfoArray)
        {
            // initialize result
            String output = "";

            // initialize text writer
            StringWriter textWriter = new StringWriter();

            
            // initialize xml serializer
            System.Xml.Serialization.XmlSerializer writerPath = new System.Xml.Serialization.XmlSerializer(typeof(CCDWrapper.DisplayConfigPathInfo));
            System.Xml.Serialization.XmlSerializer writerModeTarget = new System.Xml.Serialization.XmlSerializer(typeof(CCDWrapper.DisplayConfigTargetMode));
            System.Xml.Serialization.XmlSerializer writerModeSource = new System.Xml.Serialization.XmlSerializer(typeof(CCDWrapper.DisplayConfigSourceMode));
            System.Xml.Serialization.XmlSerializer writerModeInfoType = new System.Xml.Serialization.XmlSerializer(typeof(CCDWrapper.DisplayConfigModeInfoType));
            System.Xml.Serialization.XmlSerializer writerModeAdapterID = new System.Xml.Serialization.XmlSerializer(typeof(CCDWrapper.LUID));

            // write content to string
            textWriter.WriteLine("<displaySettings>");
            textWriter.WriteLine("<pathInfoArray>");
            foreach (CCDWrapper.DisplayConfigPathInfo pathInfo in pathInfoArray)
            {
                writerPath.Serialize(textWriter, pathInfo);
            }
            textWriter.WriteLine("</pathInfoArray>");

            textWriter.WriteLine("<modeInfoArray>");
            for (int iModeInfo = 0; iModeInfo < modeInfoArray.Length; iModeInfo++)
            {
                textWriter.WriteLine("<modeInfo>");
                CCDWrapper.DisplayConfigModeInfo modeInfo = modeInfoArray[iModeInfo];
                textWriter.WriteLine("<id>" + modeInfo.id.ToString() + "</id>");
                writerModeAdapterID.Serialize(textWriter, modeInfo.adapterId);
                writerModeInfoType.Serialize(textWriter, modeInfo.infoType);
                if (modeInfo.infoType == CCDWrapper.DisplayConfigModeInfoType.Target)
                {
                    writerModeTarget.Serialize(textWriter, modeInfo.targetMode);
                }
                else
                {
                    writerModeSource.Serialize(textWriter, modeInfo.sourceMode);
                }
                textWriter.WriteLine("</modeInfo>");
            }
            textWriter.WriteLine("</modeInfoArray>");

            output = textWriter.ToString();
            return output;
        }

        public static Boolean SaveDisplaySettings(String fileName)
        {
            CCDWrapper.DisplayConfigPathInfo[] pathInfoArray = new CCDWrapper.DisplayConfigPathInfo[0];
            CCDWrapper.DisplayConfigModeInfo[] modeInfoArray = new CCDWrapper.DisplayConfigModeInfo[0];
            CCDWrapper.MonitorAdditionalInfo[] additionalInfo = new CCDWrapper.MonitorAdditionalInfo[0];

            DebugOutput("Getting display config");
            Boolean status = GetDisplaySettings(ref pathInfoArray, ref modeInfoArray, ref additionalInfo, true);
            if (status) 
            {
                if (debug)
                {
                    // debug output complete display settings
                    DebugOutput("Display settings to write:");
                    Console.WriteLine(PrintDisplaySettings(pathInfoArray, modeInfoArray));
                }

                DebugOutput("Initializing objects for Serialization");
                System.Xml.Serialization.XmlSerializer writerAdditionalInfo = new System.Xml.Serialization.XmlSerializer(typeof(CCDWrapper.MonitorAdditionalInfo));
                System.Xml.Serialization.XmlSerializer writerPath = new System.Xml.Serialization.XmlSerializer(typeof(CCDWrapper.DisplayConfigPathInfo));
                System.Xml.Serialization.XmlSerializer writerModeTarget = new System.Xml.Serialization.XmlSerializer(typeof(CCDWrapper.DisplayConfigTargetMode));
                System.Xml.Serialization.XmlSerializer writerModeSource = new System.Xml.Serialization.XmlSerializer(typeof(CCDWrapper.DisplayConfigSourceMode));
                System.Xml.Serialization.XmlSerializer writerModeInfoType = new System.Xml.Serialization.XmlSerializer(typeof(CCDWrapper.DisplayConfigModeInfoType));
                System.Xml.Serialization.XmlSerializer writerModeAdapterID = new System.Xml.Serialization.XmlSerializer(typeof(CCDWrapper.LUID));
                XmlWriter xml = XmlWriter.Create(fileName);

                xml.WriteStartDocument();
                xml.WriteStartElement("displaySettings");
                xml.WriteStartElement("pathInfoArray");
                foreach (CCDWrapper.DisplayConfigPathInfo pathInfo in pathInfoArray)
                {
                    writerPath.Serialize(xml, pathInfo);
                }
                xml.WriteEndElement();

                xml.WriteStartElement("modeInfoArray");
                for (int iModeInfo = 0; iModeInfo < modeInfoArray.Length; iModeInfo++)
                {
                    xml.WriteStartElement("modeInfo");
                    CCDWrapper.DisplayConfigModeInfo modeInfo = modeInfoArray[iModeInfo];
                    xml.WriteElementString("id", modeInfo.id.ToString());
                    writerModeAdapterID.Serialize(xml, modeInfo.adapterId);
                    writerModeInfoType.Serialize(xml, modeInfo.infoType);
                    if (modeInfo.infoType == CCDWrapper.DisplayConfigModeInfoType.Target)
                    {
                        writerModeTarget.Serialize(xml, modeInfo.targetMode);
                    }
                    else
                    {
                        writerModeSource.Serialize(xml, modeInfo.sourceMode);
                    }
                    xml.WriteEndElement();
                }
                xml.WriteEndElement();

                xml.WriteStartElement("additionalInfo");
                for (int iAdditionalInfo = 0; iAdditionalInfo < additionalInfo.Length; iAdditionalInfo++)
                {
                    writerAdditionalInfo.Serialize(xml, additionalInfo[iAdditionalInfo]);
                }
                xml.WriteEndElement();
                xml.WriteEndDocument();
                xml.Flush();
                xml.Close();

                return true;
            }
            else
            {                    
                Console.WriteLine("Failed to get display settings, ERROR: " + status.ToString());
            }                

            return false;
        }

        static void Main(string[] args)
        {
            debug = false;
            noIDMatch = false;

            Boolean validCommand = false;
            foreach (string iArg in args)
            {
                string[] argElements = iArg.Split(new char[] { ':' }, 2);

                switch (argElements[0].ToLower())
                {
                    case "-debug":
                        debug = true;
                        DebugOutput("\nDebug output enabled");
                        break;
                    case "-noidmatch":
                        noIDMatch = true;
                        DebugOutput("\nDisabled matching of adapter IDs");
                        break;
                    case "-save":
                        SaveDisplaySettings(argElements[1]);
                        validCommand = true;
                        break;
                    case "-load":
                        LoadDisplaySettings(argElements[1]);
                        validCommand = true;
                        break;
                    case "-print":
                        CCDWrapper.DisplayConfigPathInfo[] pathInfoArray = new CCDWrapper.DisplayConfigPathInfo[0];
                        CCDWrapper.DisplayConfigModeInfo[] modeInfoArray = new CCDWrapper.DisplayConfigModeInfo[0];
                        CCDWrapper.MonitorAdditionalInfo[] additionalInfo = new CCDWrapper.MonitorAdditionalInfo[0];

                        Boolean status = GetDisplaySettings(ref pathInfoArray, ref modeInfoArray, ref additionalInfo, true);
                        if (status)
                        {
                            Console.WriteLine(PrintDisplaySettings(pathInfoArray, modeInfoArray));
                        }
                        else
                        {
                            Console.WriteLine("Failed to get display settings");
                        }
                        validCommand = true;
                        break;
                }
            }

            if (!validCommand)
            {
                Console.WriteLine("Monitor Profile Switcher command line utlility (version 0.8.0.0):\n");
                Console.WriteLine("Paremeters to MonitorSwitcher.exe:");
                Console.WriteLine("\t -save:{xmlfile} \t save the current monitor configuration to file (full path)");
                Console.WriteLine("\t -load:{xmlfile} \t load and apply monitor configuration from file (full path)");
                Console.WriteLine("\t -debug \t\t enable debug output (parameter must come before -load or -save)");
                Console.WriteLine("\t -noidmatch \t\t disable matching of adapter IDs");
                Console.WriteLine("\t -print \t\t print current monitor configuration to console");
                Console.WriteLine("");
                Console.WriteLine("Examples:");
                Console.WriteLine("\tMonitorSwitcher.exe -save:MyProfile.xml");
                Console.WriteLine("\tMonitorSwitcher.exe -load:MyProfile.xml");
                Console.WriteLine("\tMonitorSwitcher.exe -debug -load:MyProfile.xml");
                Console.ReadKey();
            }
        }
    }
}

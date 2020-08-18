/**************************************************************************
*                           MIT License
* 
* Copyright (C) 2014 Morten Kvistgaard <mk@pch-engineering.dk>
* Copyright (C) 2015 Frederic Chaxel <fchaxel@free.fr>
*
* Permission is hereby granted, free of charge, to any person obtaining
* a copy of this software and associated documentation files (the
* "Software"), to deal in the Software without restriction, including
* without limitation the rights to use, copy, modify, merge, publish,
* distribute, sublicense, and/or sell copies of the Software, and to
* permit persons to whom the Software is furnished to do so, subject to
* the following conditions:
*
* The above copyright notice and this permission notice shall be included
* in all copies or substantial portions of the Software.
*
* THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
* EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
* MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
* IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
* CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
* TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
* SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
*
*********************************************************************/

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO.BACnet;
using System.IO;
using System.IO.BACnet.Storage;
using System.Xml.Serialization;
using System.Threading;
using System.Runtime.Serialization.Formatters.Binary;
using System.Media;
using System.Linq;
using System.Collections;
using System.Reflection;
using ZedGraph;

namespace Yabe
{
    public partial class YabeMainDialog : Form
    {
        private readonly Dictionary<BacnetClient, BacnetDeviceLine> _mDevices = new Dictionary<BacnetClient, BacnetDeviceLine>();

        private readonly Dictionary<string, ListViewItem> _mSubscriptionList = new Dictionary<string, ListViewItem>();
        private readonly Dictionary<string, RollingPointPairList> _mSubscriptionPoints = new Dictionary<string, RollingPointPairList>();
        private readonly Color[] _graphColor = { Color.Red, Color.Blue, Color.Green, Color.Violet, Color.Chocolate, Color.Orange };
        private readonly GraphPane _pane;

        // Memory of all object names already discovered, first string in the Tuple is the device network address hash
        // The tuple contains two value types, so it's ok for cross session
        public Dictionary<Tuple<string, BacnetObject>, string> DevicesObjectsName = new Dictionary<Tuple<string, BacnetObject>, string>();

        private uint _mNextSubscriptionId = 0;

        private static DeviceStorage _mStorage;
        private List<BacnetObjectDescription> _objectsDescriptionExternal, _objectsDescriptionDefault;

        private YabeMainDialog _yabeFrm; // Ref to itself, already affected, useful for plugin development inside this code, before exporting it

        private class BacnetDeviceLine
        {
            public readonly BacnetClient Line;
            public readonly List<KeyValuePair<BacnetAddress, uint>> Devices = new List<KeyValuePair<BacnetAddress, uint>>();
            public readonly HashSet<byte> MstpSourcesSeen = new HashSet<byte>();
            public readonly HashSet<byte> MstpPfmDestinationsSeen = new HashSet<byte>();
            public BacnetDeviceLine(BacnetClient comm)
            {
                Line = comm;
            }
        }

        private int _asyncRequestId = 0;

        public YabeMainDialog()
        {
            _yabeFrm = this;
            
            InitializeComponent();
            Text = $"Yet Another BACnet Explorer by Igor - {Assembly.GetExecutingAssembly().GetName().Version}";
            Trace.Listeners.Add(new MyTraceListener(this));
            m_DeviceTree.ExpandAll();

            // COV Graph
            _pane = CovGraph.GraphPane;
            _pane.Title.Text = null;
            CovGraph.IsShowPointValues = true;
            // X Axis
            _pane.XAxis.Type = AxisType.Date;
            _pane.XAxis.Title.Text = null;
            _pane.XAxis.MajorGrid.IsVisible = true;
            _pane.XAxis.MajorGrid.Color = Color.Gray;
            // Y Axis
            _pane.YAxis.Title.Text = null;
            _pane.YAxis.MajorGrid.IsVisible = true;
            _pane.YAxis.MajorGrid.Color = Color.Gray;
            CovGraph.AxisChange();
            CovGraph.IsAutoScrollRange = true;

            //load splitter setup
            try
            {

                if (Properties.Settings.Default.SettingsUpgradeRequired)
                {
                    Properties.Settings.Default.Upgrade();
                    Properties.Settings.Default.SettingsUpgradeRequired = false;
                    Properties.Settings.Default.Save();
                }

                if (Properties.Settings.Default.GUI_FormSize != new Size(0, 0))
                    Size = Properties.Settings.Default.GUI_FormSize;
                var state = (FormWindowState)Enum.Parse(typeof(FormWindowState), Properties.Settings.Default.GUI_FormState);
                if (state != FormWindowState.Minimized)
                    WindowState = state;
                if (Properties.Settings.Default.GUI_SplitterButtom != -1)
                    m_SplitContainerButtom.SplitterDistance = Properties.Settings.Default.GUI_SplitterButtom;
                if (Properties.Settings.Default.GUI_SplitterLeft != -1)
                    m_SplitContainerLeft.SplitterDistance = Properties.Settings.Default.GUI_SplitterLeft;
                if (Properties.Settings.Default.GUI_SplitterRight != -1)
                    m_SplitContainerRight.SplitterDistance = Properties.Settings.Default.GUI_SplitterRight;

                // Try to open the current (if exist) object Id<-> object name mapping file
                Stream stream = File.Open(Properties.Settings.Default.ObjectNameFile, FileMode.Open);
                var bf = new BinaryFormatter();
                var d = (Dictionary<Tuple<string, BacnetObject>, string>)bf.Deserialize(stream);
                stream.Close();

                if (d != null) DevicesObjectsName = d;
            }
            catch
            {
                //ignore
            }
        }

        [Localizable(false)]
        public sealed override string Text
        {
            get { return base.Text; }
            set { base.Text = value; }
        }

        private static string ConvertToText(IList<BacnetValue> values)
        {
            if (values == null)
                return "[null]";
            else if (values.Count == 0)
                return "";
            else if (values.Count == 1)
                return values[0].Value.ToString();
            else
            {
                var ret = "{";
                foreach (var value in values)
                    ret += value.Value + ",";
                ret = ret.Substring(0, ret.Length - 1);
                ret += "}";
                return ret;
            }
        }

        private void ChangeTreeNodePropertyName(TreeNode tn, string name)
        {
            // Tooltip not set is not null, strange !
            if (tn.ToolTipText == "")
                tn.ToolTipText = tn.Text;
            if (Properties.Settings.Default.DisplayIdWithName)
                tn.Text = name + " (" + Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(tn.ToolTipText.ToLower()) + ")";
            else
                tn.Text = name;
        }

        private void SetSubscriptionStatus(ListViewItem itm, string status)
        {
            if (itm.SubItems[5].Text == status) return;
            itm.SubItems[5].Text = status;
            itm.SubItems[4].Text = DateTime.Now.ToString(Properties.Settings.Default.COVTimeFormater);
        }

        private string EventTypeNiceName(BacnetEventNotificationData.BacnetEventStates state)
        {
            return state.ToString().Substring(12);
        }


        private void OnEventNotify(BacnetClient sender, BacnetAddress adr, byte invokeId, BacnetEventNotificationData eventData, bool needConfirm)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new BacnetClient.EventNotificationCallbackHandler(OnEventNotify), new object[] { sender, adr, invokeId, eventData, needConfirm });
                return;
            }

            var subKey = eventData.initiatingObjectIdentifier.instanceId + ":" + eventData.eventObjectIdentifier.type + ":" + eventData.eventObjectIdentifier.instanceId;

            ListViewItem itm = null;
            // find the Event in the View
            foreach (ListViewItem l in m_SubscriptionView.Items)
            {
                if (l.Tag.ToString() == subKey)
                {
                    itm = l;
                    break;
                }
            }

            if (itm == null)
            {
                itm = m_SubscriptionView.Items.Add(adr.ToString());
                itm.Tag = subKey;
                itm.SubItems.Add("OBJECT_DEVICE:" + eventData.initiatingObjectIdentifier.instanceId);
                itm.SubItems.Add(eventData.eventObjectIdentifier.type + ":" + eventData.eventObjectIdentifier.instanceId);   //name
                itm.SubItems.Add(EventTypeNiceName(eventData.fromState) + " to " + EventTypeNiceName(eventData.toState));
                itm.SubItems.Add(eventData.timeStamp.Time.ToString(Properties.Settings.Default.COVTimeFormater));   //time
                itm.SubItems.Add(eventData.notifyType.ToString());   //status
            }
            else
            {
                itm.SubItems[3].Text = EventTypeNiceName(eventData.fromState) + " to " + EventTypeNiceName(eventData.toState);
                itm.SubItems[4].Text = eventData.timeStamp.Time.ToString("HH:mm:ss");   //time
                itm.SubItems[5].Text = eventData.notifyType.ToString();   //status
            }

            AddLogAlarmEvent(itm);

            //send ack
            if (needConfirm)
            {
                sender.SimpleAckResponse(adr, BacnetConfirmedServices.SERVICE_CONFIRMED_EVENT_NOTIFICATION, invokeId);
            }

        }

        private void OnCOVNotification(BacnetClient sender, BacnetAddress adr, byte invokeId, uint subscriberProcessIdentifier, BacnetObject initiatingDeviceIdentifier, BacnetObject monitoredObjectIdentifier, uint timeRemaining, bool needConfirm, ICollection<BacnetPropertyValue> values, BacnetMaxSegments maxSegments)
        {
            var subKey = adr + ":" + initiatingDeviceIdentifier.instanceId + ":" + subscriberProcessIdentifier;

            lock (_mSubscriptionList)
                if (_mSubscriptionList.ContainsKey(subKey))
                {
                    BeginInvoke((MethodInvoker)delegate
                    {
                        try
                        {
                            ListViewItem itm;
                            lock (_mSubscriptionList)
                                itm = _mSubscriptionList[subKey];

                            foreach (var value in values)
                            {

                                switch ((BacnetPropertyIds)value.property.propertyIdentifier)
                                {
                                    case BacnetPropertyIds.PROP_PRESENT_VALUE:
                                        itm.SubItems[3].Text = ConvertToText(value.value);
                                        itm.SubItems[4].Text = DateTime.Now.ToString(Properties.Settings.Default.COVTimeFormater);
                                        if (itm.SubItems[5].Text == "Not started") itm.SubItems[5].Text = "OK";

                                        try
                                        {
                                            //  try convert from string
                                            var y = Convert.ToDouble(itm.SubItems[3].Text);
                                            var x = new XDate(DateTime.Now);

                                            _pane.Title.Text = "";

                                            if ((Properties.Settings.Default.GraphLineStep) && (_mSubscriptionPoints[subKey].Count != 0))
                                            {
                                                PointPair p = _mSubscriptionPoints[subKey].Peek();
                                                _mSubscriptionPoints[subKey].Add(x, p.Y);
                                            }
                                            _mSubscriptionPoints[subKey].Add(x, y);
                                            CovGraph.AxisChange();
                                            CovGraph.Invalidate();
                                        }
                                        catch { }
                                        break;
                                    case BacnetPropertyIds.PROP_STATUS_FLAGS:
                                        if (value.value != null && value.value.Count > 0)
                                        {
                                            BacnetStatusFlags status = (BacnetStatusFlags)((BacnetBitString)value.value[0].Value).ConvertToInt();
                                            var statusText = "";
                                            if ((status & BacnetStatusFlags.STATUS_FLAG_FAULT) == BacnetStatusFlags.STATUS_FLAG_FAULT)
                                                statusText += "FAULT,";
                                            else if ((status & BacnetStatusFlags.STATUS_FLAG_IN_ALARM) == BacnetStatusFlags.STATUS_FLAG_IN_ALARM)
                                                statusText += "ALARM,";
                                            else if ((status & BacnetStatusFlags.STATUS_FLAG_OUT_OF_SERVICE) == BacnetStatusFlags.STATUS_FLAG_OUT_OF_SERVICE)
                                                statusText += "OOS,";
                                            else if ((status & BacnetStatusFlags.STATUS_FLAG_OVERRIDDEN) == BacnetStatusFlags.STATUS_FLAG_OVERRIDDEN)
                                                statusText += "OR,";
                                            if (statusText != "")
                                            {
                                                statusText = statusText.Substring(0, statusText.Length - 1);
                                                itm.SubItems[5].Text = statusText;
                                            }
                                            else
                                                itm.SubItems[5].Text = "OK";
                                        }

                                        break;
                                    default:
                                        //got something else? ignore it
                                        break;
                                }
                            }

                            AddLogAlarmEvent(itm);
                        }
                        catch (Exception ex)
                        {
                            Trace.TraceError("Exception in subscribed value: " + ex.Message);
                        }
                    });
                }

            //send ack
            if (needConfirm)
            {
                sender.SimpleAckResponse(adr, BacnetConfirmedServices.SERVICE_CONFIRMED_COV_NOTIFICATION, invokeId);
            }
        }

        #region " Trace Listner "
        private class MyTraceListener : TraceListener
        {
            private readonly YabeMainDialog _mForm;

            public MyTraceListener(YabeMainDialog form)
                : base("MyListener")
            {
                _mForm = form;
            }

            public override void TraceEvent(TraceEventCache eventCache, string source, TraceEventType eventType, int id, string message)
            {
                if ((Filter != null) && !Filter.ShouldTrace(eventCache, source, eventType, id, message, null, null, null)) return;

                ConsoleColor color;
                switch (eventType)
                {
                    case TraceEventType.Error:
                        color = ConsoleColor.Red;
                        break;
                    case TraceEventType.Warning:
                        color = ConsoleColor.Yellow;
                        break;
                    case TraceEventType.Information:
                        color = ConsoleColor.DarkGreen;
                        break;
                    default:
                        color = ConsoleColor.Gray;
                        break;
                }

                WriteColor(message + Environment.NewLine, color);
            }

            public override void TraceEvent(TraceEventCache eventCache, string source, TraceEventType eventType, int id, string format, params object[] args)
            {
                if ((Filter != null) && !Filter.ShouldTrace(eventCache, source, eventType, id, format, args, null, null)) return;

                ConsoleColor color;
                switch (eventType)
                {
                    case TraceEventType.Error:
                        color = ConsoleColor.Red;
                        break;
                    case TraceEventType.Warning:
                        color = ConsoleColor.Yellow;
                        break;
                    case TraceEventType.Information:
                        color = ConsoleColor.DarkGreen;
                        break;
                    default:
                        color = ConsoleColor.Gray;
                        break;
                }

                WriteColor(string.Format(format, args) + Environment.NewLine, color);
            }

            public override void Write(string message)
            {
                WriteColor(message, ConsoleColor.Gray);
            }
            public override void WriteLine(string message)
            {
                WriteColor(message + Environment.NewLine, ConsoleColor.Gray);
            }

            private void WriteColor(string message, ConsoleColor color)
            {
                if (!_mForm.IsHandleCreated) return;

                _mForm.m_LogText.BeginInvoke((MethodInvoker)delegate { _mForm.m_LogText.AppendText(message); });
            }
        }
        #endregion

        private void MainDialog_Load(object sender, EventArgs e)
        {
            //start renew timer at half lifetime
            var lifetime = (int)Properties.Settings.Default.Subscriptions_Lifetime;
            if (lifetime > 0)
            {
                m_subscriptionRenewTimer.Interval = (lifetime / 2) * 1000;
                m_subscriptionRenewTimer.Enabled = true;
            }

            //display nice floats in propertyGrid
            Utilities.CustomSingleConverter.DontDisplayExactFloats = true;

            m_DeviceTree.TreeViewNodeSorter = new NodeSorter();

            var listPlugins = Properties.Settings.Default.Plugins.Split(new char[] { ',', ';' });

            foreach (var pluginName in listPlugins)
            {
                try
                {
                    var path = Path.GetDirectoryName(Application.ExecutablePath);
                    var name = pluginName.Replace(" ", string.Empty);
                    var myDll = Assembly.LoadFrom(path + "/" + name + ".dll");
                    var types = myDll.GetExportedTypes();
                    var plugin = (IYabePlugin)myDll.CreateInstance(name + ".Plugin", true);
                    plugin?.Init(this);
                }
                catch
                {
                    Trace.WriteLine("Error loading plugins " + pluginName);
                }
            }

            if (pluginsToolStripMenuItem.DropDownItems.Count == 0) pluginsToolStripMenuItem.Visible = false;
        }

        private TreeNode FindDeviceTreeNode(BacnetClient comm)
        {
            foreach (TreeNode node in m_DeviceTree.Nodes[0].Nodes)
            {
                if (node.Tag is BacnetClient c && c.Equals(comm)) return node;
            }
            return null;
        }

        private TreeNode FindDeviceTreeNode(IBacnetTransport transport)
        {
            foreach (TreeNode node in m_DeviceTree.Nodes[0].Nodes)
            {
                if (node.Tag is BacnetClient c && c.Transport.Equals(transport)) return node;
            }
            return null;
        }

        // Only the see Yabe on the net
        private static void OnWhoIs(BacnetClient sender, BacnetAddress adr, int lowLimit, int highLimit)
        {
            var myId = (uint)Properties.Settings.Default.YabeDeviceId;

            if (lowLimit != -1 && myId < lowLimit) return;
            else if (highLimit != -1 && myId > highLimit) return;
            sender.Iam(myId, BacnetSegmentations.SEGMENTATION_BOTH, 61440);
        }

        private static void OnWhoIsIgnore(BacnetClient sender, BacnetAddress adr, int lowLimit, int highLimit)
        {
            //ignore whoIs responses from other devices (or loopbacks)
        }

        private static void OnReadPropertyRequest(BacnetClient sender, BacnetAddress adr, byte invokeId, BacnetObject objectId, BacnetPropertyReference property, BacnetMaxSegments maxSegments)
        {
            lock (_mStorage)
            {
                try
                {
                    var code = _mStorage.ReadProperty(objectId, (BacnetPropertyIds)property.propertyIdentifier, property.propertyArrayIndex, out var value);
                    if (code == DeviceStorage.ErrorCodes.Good)
                        sender.ReadPropertyResponse(adr, invokeId, sender.GetSegmentBuffer(maxSegments), objectId, property, value);
                    else
                        sender.ErrorResponse(adr, BacnetConfirmedServices.SERVICE_CONFIRMED_READ_PROPERTY, invokeId, BacnetErrorClasses.ERROR_CLASS_DEVICE, BacnetErrorCodes.ERROR_CODE_OTHER);
                }
                catch (Exception)
                {
                    sender.ErrorResponse(adr, BacnetConfirmedServices.SERVICE_CONFIRMED_READ_PROPERTY, invokeId, BacnetErrorClasses.ERROR_CLASS_DEVICE, BacnetErrorCodes.ERROR_CODE_OTHER);
                }
            }
        }
        private static void OnReadPropertyMultipleRequest(BacnetClient sender, BacnetAddress adr, byte invokeId, IList<BacnetReadAccessSpecification> properties, BacnetMaxSegments maxSegments)
        {
            lock (_mStorage)
            {
                try
                {
                    var values = new List<BacnetReadAccessResult>();
                    foreach (var p in properties)
                    {
                        IList<BacnetPropertyValue> value;
                        if (p.propertyReferences.Count == 1 && p.propertyReferences[0].propertyIdentifier == (uint)BacnetPropertyIds.PROP_ALL)
                        {
                            if (!_mStorage.ReadPropertyAll(p.objectIdentifier, out value))
                            {
                                sender.ErrorResponse(adr, BacnetConfirmedServices.SERVICE_CONFIRMED_READ_PROP_MULTIPLE, invokeId, BacnetErrorClasses.ERROR_CLASS_OBJECT, BacnetErrorCodes.ERROR_CODE_UNKNOWN_OBJECT);
                                return;
                            }
                        }
                        else
                            _mStorage.ReadPropertyMultiple(p.objectIdentifier, p.propertyReferences, out value);
                        values.Add(new BacnetReadAccessResult(p.objectIdentifier, value));
                    }

                    sender.ReadPropertyMultipleResponse(adr, invokeId, sender.GetSegmentBuffer(maxSegments), values);

                }
                catch (Exception)
                {
                    sender.ErrorResponse(adr, BacnetConfirmedServices.SERVICE_CONFIRMED_READ_PROP_MULTIPLE, invokeId, BacnetErrorClasses.ERROR_CLASS_DEVICE, BacnetErrorCodes.ERROR_CODE_OTHER);
                }
            }
        }

        private void OnIam(BacnetClient sender, BacnetAddress adr, uint deviceId, uint maxApdu, BacnetSegmentations segmentation, ushort vendorId)
        {
            var newEntry = new KeyValuePair<BacnetAddress, uint>(adr, deviceId);
            if (!_mDevices.ContainsKey(sender)) return;
            if (!_mDevices[sender].Devices.Contains(newEntry))
                _mDevices[sender].Devices.Add(newEntry);
            else
                return;

            //update GUI
            BeginInvoke((MethodInvoker)delegate
            {
                var parent = FindDeviceTreeNode(sender);
                if (parent == null) return;

                var propObjectNameOk = false;
                string identifier = null;

                lock (DevicesObjectsName)
                    propObjectNameOk = DevicesObjectsName.TryGetValue(new Tuple<string, BacnetObject>(adr.FullHashString(), new BacnetObject(BacnetObjectTypes.OBJECT_DEVICE, deviceId)), out identifier);

                //update existing (this can happen in MSTP)
                foreach (TreeNode s in parent.Nodes)
                {
                    if (s.Tag is KeyValuePair<BacnetAddress, uint> entry && entry.Key.Equals(adr))
                    {
                        s.Text = "Device " + newEntry.Value + " - " + newEntry.Key.ToString(s.Parent.Parent != null);
                        s.Tag = newEntry;
                        if (propObjectNameOk)
                        {
                            s.ToolTipText = s.Text;
                            s.Text = identifier + " [" + deviceId + "] ";
                        }

                        return;
                    }
                }
                // Try to add it under a router if any 
                foreach (TreeNode s in parent.Nodes)
                {
                    if (s.Tag is KeyValuePair<BacnetAddress, uint> entry && entry.Key.IsMyRouter(adr))
                    {
                        var node =
                            new TreeNode("Device " + newEntry.Value + " - " + newEntry.Key.ToString(true))
                            {
                                ImageIndex = 2
                            };
                        node.SelectedImageIndex = node.ImageIndex;
                        node.Tag = newEntry;
                        if (propObjectNameOk)
                        {
                            node.ToolTipText = node.Text;
                            node.Text = identifier + " [" + deviceId + "] ";
                        }
                        s.Nodes.Add(node);
                        m_DeviceTree.ExpandAll();
                        return;
                    }
                }

                //add simply
                var basicNode =
                    new TreeNode("Device " + newEntry.Value + " - " + newEntry.Key.ToString(false)) {ImageIndex = 2};
                basicNode.SelectedImageIndex = basicNode.ImageIndex;
                basicNode.Tag = newEntry;
                if (propObjectNameOk)
                {
                    basicNode.ToolTipText = basicNode.Text;
                    basicNode.Text = identifier + " [" + deviceId + "] ";
                }
                parent.Nodes.Add(basicNode);
                m_DeviceTree.ExpandAll();
            });
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show(this, "Yet Another Bacnet Explorer - Igor\nVersion " + GetType().Assembly.GetName().Version + "\nBy Morten Kvistgaard - Copyright 2014-2017\nBy Frederic Chaxel - Copyright 2015-2018\nBy Chris Jenson - Copyright 2020\n" +
                "\nReferences:" +
                "\nhttp://bacnet.sourceforge.net/" +
                "\nhttp://www.unified-automation.com/products/development-tools/uaexpert.html" +
                "\nhttp://www.famfamfam.com/" +
                "\nhttp://sourceforge.net/projects/zedgraph/" +
                "\nhttp://www.codeproject.com/Articles/38699/A-Professional-Calendar-Agenda-View-That-You-Will" +
                "\nhttps://github.com/chmorgan/sharppcap" +
                "\nhttps://sourceforge.net/projects/mstreeview"

                , "About", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void quitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void addDeviceSearchToolStripMenuItem_Click(object sender, EventArgs e)
        {
            labelDrop1.Visible = labelDrop2.Visible = false;

            var dlg = new SearchDialog();
            if (dlg.ShowDialog(this) == DialogResult.OK)
            {
                var comm = dlg.Result;
                try
                {
                    _mDevices.Add(comm, new BacnetDeviceLine(comm));
                }
                catch { return; }

                //add to tree
                var node = m_DeviceTree.Nodes[0].Nodes.Add(comm.ToString());
                node.Tag = comm;
                switch (comm.Transport.Type)
                {
                    case BacnetAddressTypes.IP:
                        node.ImageIndex = 3;
                        break;
                    case BacnetAddressTypes.MSTP:
                        node.ImageIndex = 1;
                        break;
                    default:
                        node.ImageIndex = 8;
                        break;
                }
                node.SelectedImageIndex = node.ImageIndex;
                m_DeviceTree.ExpandAll(); m_DeviceTree.SelectedNode = node;

                try
                {
                    //start BACnet
                    comm.ProposedWindowSize = Properties.Settings.Default.Segments_ProposedWindowSize;
                    comm.Retries = (int)Properties.Settings.Default.DefaultRetries;
                    comm.Timeout = (int)Properties.Settings.Default.DefaultTimeout;
                    comm.MaxSegments = BacnetClient.GetSegmentsCount(Properties.Settings.Default.Segments_Max);
                    if (Properties.Settings.Default.YabeDeviceId >= 0) // If Yabe get a Device id
                    {
                        if (_mStorage == null)
                        {
                            // Load descriptor from the embedded xml resource
                            _mStorage = _mStorage = DeviceStorage.Load("Yabe.YabeDeviceDescriptor.xml", (uint)Properties.Settings.Default.YabeDeviceId);
                            // A fast way to change the PROP_OBJECT_LIST
                            var prop = Array.Find<Property>(_mStorage.Objects[0].Properties, p => p.Id == BacnetPropertyIds.PROP_OBJECT_LIST);
                            prop.Value[0] = "OBJECT_DEVICE:" + Properties.Settings.Default.YabeDeviceId;
                            // change PROP_FIRMWARE_REVISION
                            prop = Array.Find<Property>(_mStorage.Objects[0].Properties, p => p.Id == BacnetPropertyIds.PROP_FIRMWARE_REVISION);
                            prop.Value[0] = GetType().Assembly.GetName().Version.ToString();
                            // change PROP_APPLICATION_SOFTWARE_VERSION
                            prop = Array.Find<Property>(_mStorage.Objects[0].Properties, p => p.Id == BacnetPropertyIds.PROP_APPLICATION_SOFTWARE_VERSION);
                            prop.Value[0] = GetType().Assembly.GetName().Version.ToString();
                        }
                        comm.OnWhoIs += new BacnetClient.WhoIsHandler(OnWhoIs);
                        comm.OnReadPropertyRequest += new BacnetClient.ReadPropertyRequestHandler(OnReadPropertyRequest);
                        comm.OnReadPropertyMultipleRequest += new BacnetClient.ReadPropertyMultipleRequestHandler(OnReadPropertyMultipleRequest);
                    }
                    else
                    {
                        comm.OnWhoIs += new BacnetClient.WhoIsHandler(OnWhoIsIgnore);
                    }
                    comm.OnIam += new BacnetClient.IamHandler(OnIam);
                    comm.OnCOVNotification += new BacnetClient.COVNotificationHandler(OnCOVNotification);
                    comm.OnEventNotify += new BacnetClient.EventNotificationCallbackHandler(OnEventNotify);
                    comm.Start();

                    //start search
                    if (comm.Transport.Type == BacnetAddressTypes.IP || comm.Transport.Type == BacnetAddressTypes.Ethernet
                        || comm.Transport.Type == BacnetAddressTypes.IPV6
                        || (comm.Transport is BacnetMstpProtocolTransport && ((BacnetMstpProtocolTransport)comm.Transport).SourceAddress != -1)
                        || comm.Transport.Type == BacnetAddressTypes.PTP)
                    {
                        ThreadPool.QueueUserWorkItem((o) =>
                        {
                            for (var i = 0; i < comm.Retries; i++)
                            {
                                comm.WhoIs();
                                Thread.Sleep(comm.Timeout);
                            }
                        }, null);
                    }

                    //special MSTP auto discovery
                    if (comm.Transport is BacnetMstpProtocolTransport)
                    {
                        ((BacnetMstpProtocolTransport)comm.Transport).FrameRecieved += new BacnetMstpProtocolTransport.FrameRecievedHandler(MSTP_FrameReceived);
                    }
                }
                catch (Exception ex)
                {
                    _mDevices.Remove(comm);
                    node.Remove();
                    MessageBox.Show(this, "Couldn't start Bacnet communication: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void MSTP_FrameReceived(BacnetMstpProtocolTransport sender, BacnetMstpFrameTypes frameType, byte destinationAddress, byte sourceAddress, int msgLength)
        {
            try
            {
                if (IsDisposed) return;
                BacnetDeviceLine deviceLine = null;
                foreach (var l in _mDevices.Values)
                {
                    if (l.Line.Transport == sender)
                    {
                        deviceLine = l;
                        break;
                    }
                }
                if (deviceLine == null) return;
                lock (deviceLine.MstpSourcesSeen)
                {
                    if (!deviceLine.MstpSourcesSeen.Contains(sourceAddress))
                    {
                        deviceLine.MstpSourcesSeen.Add(sourceAddress);

                        //find parent node
                        var parent = FindDeviceTreeNode(sender);

                        //find "free" node. The "free" node might have been added
                        TreeNode freeNode = null;
                        foreach (TreeNode n in parent.Nodes)
                        {
                            if (n.Text == "free" + sourceAddress)
                            {
                                freeNode = n;
                                break;
                            }
                        }

                        //update gui
                        Invoke((MethodInvoker)delegate
                        {
                            var node = parent.Nodes.Add("device" + sourceAddress);
                            node.ImageIndex = 2;
                            node.SelectedImageIndex = node.ImageIndex;
                            node.Tag = new KeyValuePair<BacnetAddress, uint>(new BacnetAddress(BacnetAddressTypes.MSTP, 0, new byte[] { sourceAddress }), 0xFFFFFFFF);
                            if (freeNode != null) freeNode.Remove();
                            m_DeviceTree.ExpandAll();
                        });

                        //detect collision
                        if (sourceAddress == sender.SourceAddress)
                        {
                            Invoke((MethodInvoker)delegate
                            {
                                MessageBox.Show(this, "Selected source address seems to be occupied!", "Collision detected", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            });
                        }
                    }
                    if (frameType == BacnetMstpFrameTypes.FRAME_TYPE_POLL_FOR_MASTER && !deviceLine.MstpPfmDestinationsSeen.Contains(destinationAddress) && sender.SourceAddress != destinationAddress)
                    {
                        deviceLine.MstpPfmDestinationsSeen.Add(destinationAddress);
                        if (!deviceLine.MstpSourcesSeen.Contains(destinationAddress) && Properties.Settings.Default.MSTP_DisplayFreeAddresses)
                        {
                            var parent = FindDeviceTreeNode(sender);
                            if (IsDisposed) return;
                            Invoke((MethodInvoker)delegate
                            {
                                var node = parent.Nodes.Add("free" + destinationAddress);
                                node.ImageIndex = 9;
                                node.SelectedImageIndex = node.ImageIndex;
                                m_DeviceTree.ExpandAll();
                            });
                        }
                    }
                }
            }
            catch (ObjectDisposedException)
            {
                //we're closing down ... ignore
            }
        }

        private void m_SearchToolButton_Click(object sender, EventArgs e)
        {
            addDeviceSearchToolStripMenuItem_Click(this, null);
        }

        private void RemoveSubscriptions(BacnetAddress adr, uint deviceId, BacnetClient comm)
        {
            var deletes = new LinkedList<string>();
            foreach (var entry in _mSubscriptionList)
            {
                var sub = (Subscription)entry.Value.Tag;
                if (((sub.Address == adr) && (sub.DeviceId.instanceId == deviceId)) || (sub.Client == comm))
                {
                    m_SubscriptionView.Items.Remove(entry.Value);
                    deletes.AddLast(sub.SubscriptionKey);
                }
            }
            foreach (var subKey in deletes)
            {
                _mSubscriptionList.Remove(subKey);
                try
                {
                    var points = _mSubscriptionPoints[subKey];
                    foreach (LineItem l in _pane.CurveList)
                        if (l.Tag == points)
                        {
                            _pane.CurveList.Remove(l);
                            break;
                        }

                    _mSubscriptionPoints.Remove(subKey);
                }
                catch { }
            }

            CovGraph.AxisChange();
            CovGraph.Invalidate();
        }

        private void removeDeviceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (m_DeviceTree.SelectedNode == null) return;
            else if (m_DeviceTree.SelectedNode.Tag == null) return;
            var deviceEntry = m_DeviceTree.SelectedNode.Tag as KeyValuePair<BacnetAddress, uint>?;
            BacnetClient commEntry;
            if (m_DeviceTree.SelectedNode.Tag is BacnetClient)
                commEntry = m_DeviceTree.SelectedNode.Tag as BacnetClient;
            else
                commEntry = m_DeviceTree.SelectedNode.Parent.Tag as BacnetClient;

            if (deviceEntry != null)
            {
                if (MessageBox.Show(this, "Delete this device?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    BacnetClient comm;
                    if (m_DeviceTree.SelectedNode.Parent.Tag is BacnetClient)
                        comm = m_DeviceTree.SelectedNode.Parent.Tag as BacnetClient;
                    else
                        comm = m_DeviceTree.SelectedNode.Parent.Parent.Tag as BacnetClient; // device under a router

                    _mDevices[comm].Devices.Remove((KeyValuePair<BacnetAddress, uint>)deviceEntry);

                    m_DeviceTree.Nodes.Remove(m_DeviceTree.SelectedNode);
                    RemoveSubscriptions(deviceEntry.Value.Key, deviceEntry.Value.Value, null);
                }
            }
            else if (commEntry != null)
            {
                if (MessageBox.Show(this, "Delete this transport?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    _mDevices.Remove(commEntry);
                    m_DeviceTree.Nodes.Remove(m_DeviceTree.SelectedNode);
                    RemoveSubscriptions(null, 0, commEntry);
                    commEntry.Dispose();
                }
            }
        }

        private void m_RemoveToolButton_Click(object sender, EventArgs e)
        {
            removeDeviceToolStripMenuItem_Click(this, null);
        }

        public static int GetIconNum(BacnetObjectTypes objectType)
        {
            switch (objectType)
            {
                case BacnetObjectTypes.OBJECT_DEVICE:
                    return 2;
                case BacnetObjectTypes.OBJECT_FILE:
                    return 5;
                case BacnetObjectTypes.OBJECT_ANALOG_INPUT:
                case BacnetObjectTypes.OBJECT_ANALOG_OUTPUT:
                case BacnetObjectTypes.OBJECT_ANALOG_VALUE:
                    return 6;
                case BacnetObjectTypes.OBJECT_BINARY_INPUT:
                case BacnetObjectTypes.OBJECT_BINARY_OUTPUT:
                case BacnetObjectTypes.OBJECT_BINARY_VALUE:
                    return 7;
                case BacnetObjectTypes.OBJECT_GROUP:
                    return 10;
                case BacnetObjectTypes.OBJECT_STRUCTURED_VIEW:
                    return 11;
                case BacnetObjectTypes.OBJECT_TRENDLOG:
                    return 12;
                case BacnetObjectTypes.OBJECT_TREND_LOG_MULTIPLE:
                    return 12;
                case BacnetObjectTypes.OBJECT_NOTIFICATION_CLASS:
                    return 13;
                case BacnetObjectTypes.OBJECT_SCHEDULE:
                    return 14;
                case BacnetObjectTypes.OBJECT_CALENDAR:
                    return 15;
                default:
                    return 4;
            }
        }
        private void SetNodeIcon(BacnetObjectTypes objectType, TreeNode node)
        {
            node.ImageIndex = GetIconNum(objectType);
            node.SelectedImageIndex = node.ImageIndex;
        }

        private void AddObjectEntry(BacnetClient client, BacnetAddress address, string name, BacnetObject bacnetObject, TreeNodeCollection nodes)
        {
            if (string.IsNullOrEmpty(name)) name = bacnetObject.ToString();

            TreeNode node;

            if (name.StartsWith("OBJECT_"))
                node = nodes.Add(name.Substring(7));
            else
                node = nodes.Add("PROPRIETARY:" + bacnetObject.Instance + " (" + name + ")");  // Proprietary Objects not in enum appears only with the number such as 584:0

            node.Tag = bacnetObject;
            node.Checked = bacnetObject.Exportable;

            //icon
            SetNodeIcon(bacnetObject.type, node);

            // Get the property name if already known
            string propertyName;

            lock (DevicesObjectsName)
                if (DevicesObjectsName.TryGetValue(new Tuple<string, BacnetObject>(address.FullHashString(), bacnetObject), out propertyName) == true)
                {
                    ChangeTreeNodePropertyName(node, propertyName); ;
                }

            //fetch sub properties
            if (bacnetObject.type == BacnetObjectTypes.OBJECT_GROUP)
                FetchGroupProperties(client, address, bacnetObject, node.Nodes);
            else if ((bacnetObject.type == BacnetObjectTypes.OBJECT_STRUCTURED_VIEW) && Properties.Settings.Default.DefaultPreferStructuredView)
                FetchViewObjects(client, address, bacnetObject, node.Nodes);
        }

        private static IList<BacnetValue> FetchStructuredObjects(BacnetClient comm, BacnetAddress adr, uint deviceId)
        {
            IList<BacnetValue> ret;
            var oldReties = comm.Retries;
            try
            {
                comm.Retries = 1;       //only do 1 retry
                if (!comm.ReadPropertyRequest(adr, new BacnetObject(BacnetObjectTypes.OBJECT_DEVICE, deviceId), BacnetPropertyIds.PROP_STRUCTURED_OBJECT_LIST, out ret))
                {
                    Trace.TraceInformation("Didn't get response from 'Structured Object List'");
                    return null;
                }
                return ret == null || ret.Count == 0 ? null : ret;
            }
            catch (Exception)
            {
                Trace.TraceInformation("Got exception from 'Structured Object List'");
                return null;
            }
            finally
            {
                comm.Retries = oldReties;
            }
        }

        private void AddObjectListOneByOneAsync(BacnetClient comm, BacnetAddress adr, uint deviceId, uint count, int asynchRequestId)
        {
            ThreadPool.QueueUserWorkItem((o) =>
            {
                try
                {
                    for (var i = 1; i <= count; i++)
                    {
                        if (!comm.ReadPropertyRequest(adr, new BacnetObject(BacnetObjectTypes.OBJECT_DEVICE, deviceId), BacnetPropertyIds.PROP_OBJECT_LIST, out var valueList, 0, (uint)i))
                        {
                            MessageBox.Show("Couldn't fetch object list index", "Communication Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                        if (asynchRequestId != _asyncRequestId) return; // Selected device is no more the good one

                        //add to tree
                        foreach (var value in valueList)
                        {
                            Invoke((MethodInvoker)delegate
                            {
                                if (asynchRequestId != _asyncRequestId) return;  // another test in the GUI thread
                                AddObjectEntry(comm, adr, null, (BacnetObject)value.Value, m_AddressSpaceTree.Nodes);
                            });
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error during read: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            });
        }

        private static List<BacnetObject> SortBacnetObjects(IList<BacnetValue> rawList)
        {

            var sortedList = new List<BacnetObject>();
            foreach (var value in rawList)
                if (value.Value is BacnetObject) // with BacnetObjectId
                    sortedList.Add((BacnetObject)value.Value);
                else // with Subordinate_List for StructuredView
                {
                    var v = (BacnetDeviceObjectReference)value.Value;
                    sortedList.Add(v.objectIdentifier); // ignore deviceIdentifier
                }

            sortedList.Sort();

            return sortedList;
        }
        private void m_AddressSpaceTree_AfterCheck(object sender, TreeViewEventArgs e)
        {
            if (!(e.Node.Tag is BacnetObject bacnetObject)) return;

            bacnetObject.Exportable = e.Node.Checked;
        }
        private void m_DeviceTree_BeforeCheck(object sender, TreeViewCancelEventArgs e)
        {
            _asyncRequestId++; // disabled a possible thread pool work (update) on the AddressSpaceTree
            if (!(e.Node.Tag is KeyValuePair<BacnetAddress, uint> entry))
            {
                e.Cancel = true;
            }
        }

        private void m_DeviceTree_AfterCheck(object sender, TreeViewEventArgs e)
        {
            _asyncRequestId++; // disabled a possible thread pool work (update) on the AddressSpaceTree
            if (_busy) return;
            _busy = true;
            try
            {
                CheckChildNodes(e.Node, e.Node.Checked);
                CheckObjectNodes(e.Node.Tag, e.Node.Checked);
            }
            finally
            {
                _busy = false;
            }
        }

        private static void CheckObjectNodes(object tag, bool @checked)
        {
            if (!(tag is KeyValuePair<BacnetAddress, uint> device)) return;
            {
                var deviceId = device.Value;
            }
        }

        private static void CheckChildNodes(TreeNode node, bool @checked)
        {
            foreach (TreeNode child in node.Nodes)
            {
                child.Checked = @checked;
                CheckChildNodes(child, @checked);
            }
        }

        private void m_DeviceTree_AfterSelect(object sender, TreeViewEventArgs e)
        {
            _asyncRequestId++; // disabled a possible thread pool work (update) on the AddressSpaceTree

            if (e.Node.Tag is KeyValuePair<BacnetAddress, uint> entry)
            {
                m_AddressSpaceTree.Nodes.Clear();   //clear
                AddSpaceLabel.Text = "Address Space";

                BacnetClient client;

                if (e.Node.Parent.Tag is BacnetClient)  // A 'basic node'
                    client = (BacnetClient)e.Node.Parent.Tag;
                else  // A routed node
                    client = (BacnetClient)e.Node.Parent.Parent.Tag;

                var bacnetAddress = entry.Key;
                var deviceId = entry.Value;

                //unconfigured MSTP?
                if (client.Transport is BacnetMstpProtocolTransport && ((BacnetMstpProtocolTransport)client.Transport).SourceAddress == -1)
                {
                    if (MessageBox.Show("The MSTP transport is not yet configured. Would you like to set source_address now?", "Set Source Address", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;

                    //find suggested address
                    byte address = 0xFF;
                    var line = _mDevices[client];
                    lock (line.MstpSourcesSeen)
                    {
                        foreach (var s in line.MstpPfmDestinationsSeen)
                        {
                            if (s < address && !line.MstpSourcesSeen.Contains(s))
                                address = s;
                        }
                    }

                    //display choice
                    var dlg = new SourceAddressDialog {SourceAddress = address};
                    if (dlg.ShowDialog(this) == DialogResult.Cancel) return;
                    ((BacnetMstpProtocolTransport)client.Transport).SourceAddress = dlg.SourceAddress;
                    Application.DoEvents();     //let the interface relax
                }

                //update "address space"?
                Cursor = Cursors.WaitCursor;
                Application.DoEvents();
                var oldTimeout = client.Timeout;
                IList<BacnetValue> valueList = null;
                try
                {
                    //fetch structured view if possible
                    if (Properties.Settings.Default.DefaultPreferStructuredView)
                        valueList = FetchStructuredObjects(client, bacnetAddress, deviceId);

                    //fetch normal list
                    if (valueList == null)
                    {
                        try
                        {
                            if (!client.ReadPropertyRequest(bacnetAddress, new BacnetObject(BacnetObjectTypes.OBJECT_DEVICE, deviceId), BacnetPropertyIds.PROP_OBJECT_LIST, out valueList))
                            {
                                Trace.TraceWarning("Didn't get response from 'Object List'");
                                valueList = null;
                            }
                        }
                        catch (Exception)
                        {
                            Trace.TraceWarning("Got exception from 'Object List'");
                            valueList = null;
                        }
                    }

                    //fetch list one-by-one
                    if (valueList == null)
                    {
                        try
                        {
                            //fetch object list count
                            if (!client.ReadPropertyRequest(bacnetAddress, new BacnetObject(BacnetObjectTypes.OBJECT_DEVICE, deviceId), BacnetPropertyIds.PROP_OBJECT_LIST, out valueList, 0, 0))
                            {
                                MessageBox.Show(this, "Couldn't fetch objects", "Communication Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(this, "Error during read: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        if (valueList != null && valueList.Count == 1 && valueList[0].Value is uint)
                        {
                            var listCount = (uint)valueList[0].Value;
                            AddSpaceLabel.Text = "Address Space : " + listCount + " objects";
                            AddObjectListOneByOneAsync(client, bacnetAddress, deviceId, listCount, _asyncRequestId);
                            return;
                        }
                        else
                        {
                            MessageBox.Show(this, "Couldn't read 'Object List' count", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }

                    var objectList = SortBacnetObjects(valueList);
                    AddSpaceLabel.Text = "Address Space : " + objectList.Count + " objects";
                    //add to tree
                    foreach (var bacnetObject in objectList)
                    {
                        // Add FC
                        // If the Device name not set, try to update it
                        if (bacnetObject.type == BacnetObjectTypes.OBJECT_DEVICE)
                        {
                            if (e.Node.ToolTipText == "")   // already update with the device name
                            {
                                var propObjectNameOk = false;
                                string identifier;

                                lock (DevicesObjectsName)
                                    propObjectNameOk = DevicesObjectsName.TryGetValue(new Tuple<string, BacnetObject>(bacnetAddress.FullHashString(), bacnetObject), out identifier);
                                if (propObjectNameOk)
                                {
                                    e.Node.ToolTipText = e.Node.Text;
                                    e.Node.Text = identifier + " [" + bacnetObject.Instance + "] ";
                                }
                                else
                                    try
                                    {
                                        IList<BacnetValue> values;
                                        if (client.ReadPropertyRequest(bacnetAddress, bacnetObject, BacnetPropertyIds.PROP_OBJECT_NAME, out values))
                                        {
                                            e.Node.ToolTipText = e.Node.Text;   // IP or MSTP node id -> in the Tooltip
                                            e.Node.Text = values[0] + " [" + bacnetObject.Instance + "] ";  // change @ by the Name    
                                            lock (DevicesObjectsName)
                                            {
                                                var t = new Tuple<string, BacnetObject>(bacnetAddress.FullHashString(), bacnetObject);
                                                DevicesObjectsName.Remove(t);
                                                DevicesObjectsName.Add(t, values[0].ToString());
                                            }
                                        }
                                    }
                                    catch { }
                            }
                        }

                        AddObjectEntry(client, bacnetAddress, null, bacnetObject, m_AddressSpaceTree.Nodes);//AddObjectEntry(comm, adr, null, bobj_id, e.Node.Nodes); 
                    }
                }
                finally
                {
                    Cursor = Cursors.Default;
                    m_DataGrid.SelectedObject = null;
                }
            }
        }

        private void FetchViewObjects(BacnetClient comm, BacnetAddress adr, BacnetObject objectId, TreeNodeCollection nodes)
        {
            try
            {
                IList<BacnetValue> values;
                if (comm.ReadPropertyRequest(adr, objectId, BacnetPropertyIds.PROP_SUBORDINATE_LIST, out values))
                {
                    var objectList = SortBacnetObjects(values);
                    foreach (var objId in objectList)
                        AddObjectEntry(comm, adr, null, objId, nodes);
                }
                else
                {
                    Trace.TraceWarning("Couldn't fetch view members");
                }
            }
            catch (Exception ex)
            {
                Trace.TraceError("Couldn't fetch view members: " + ex.Message);
            }
        }

        private void FetchGroupProperties(BacnetClient comm, BacnetAddress adr, BacnetObject objectId, TreeNodeCollection nodes)
        {
            try
            {
                IList<BacnetValue> values;
                if (comm.ReadPropertyRequest(adr, objectId, BacnetPropertyIds.PROP_LIST_OF_GROUP_MEMBERS, out values))
                {
                    foreach (var value in values)
                    {
                        if (value.Value is BacnetReadAccessSpecification)
                        {
                            var spec = (BacnetReadAccessSpecification)value.Value;
                            foreach (BacnetPropertyReference p in spec.propertyReferences)
                            {
                                AddObjectEntry(comm, adr, spec.objectIdentifier + ":" + ((BacnetPropertyIds)p.propertyIdentifier), spec.objectIdentifier, nodes);
                            }
                        }
                    }
                }
                else
                {
                    Trace.TraceWarning("Couldn't fetch group members");
                }
            }
            catch (Exception ex)
            {
                Trace.TraceError("Couldn't fetch group members: " + ex.Message);
            }
        }

        private void addDeviceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            addDeviceSearchToolStripMenuItem_Click(this, null);
        }

        private void removeDeviceToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            removeDeviceToolStripMenuItem_Click(this, null);
        }

        private static string GetNiceName(BacnetPropertyIds property)
        {
            var name = property.ToString();
            if (name.StartsWith("PROP_"))
            {
                name = name.Substring(5);
                name = name.Replace('_', ' ');
                name = Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(name.ToLower());
            }
            else
                //name = "Proprietary (" + property.ToString() + ")";
                name = property + " - Proprietary";
            return name;
        }

        private bool ReadProperty(BacnetClient comm, BacnetAddress adr, BacnetObject objectId, BacnetPropertyIds propertyId, ref IList<BacnetPropertyValue> values, uint arrayIndex = System.IO.BACnet.Serialize.ASN1.BACNET_ARRAY_ALL)
        {
            var newEntry = new BacnetPropertyValue
            {
                property = new BacnetPropertyReference((uint) propertyId, arrayIndex)
            };
            IList<BacnetValue> value;
            try
            {
                if (!comm.ReadPropertyRequest(adr, objectId, propertyId, out value, 0, arrayIndex))
                    return false;     //ignore
            }
            catch
            {
                return false;         //ignore
            }
            newEntry.value = value;
            values.Add(newEntry);
            return true;
        }

        private bool ReadAllPropertiesBySingle(BacnetClient comm, BacnetAddress adr, BacnetObject objectId, out IList<BacnetReadAccessResult> valueList)
        {

            if (_objectsDescriptionDefault == null)  // first call, Read Objects description from internal & optional external xml file
            {
                StreamReader sr;
                var xs = new XmlSerializer(typeof(List<BacnetObjectDescription>));

                // embedded resource
                Assembly assembly;
                assembly = Assembly.GetExecutingAssembly();
                sr = new StreamReader(assembly.GetManifestResourceStream("Yabe.ReadSinglePropDescrDefault.xml"));
                _objectsDescriptionDefault = (List<BacnetObjectDescription>)xs.Deserialize(sr);

                try  // External optional file
                {
                    sr = new StreamReader("ReadSinglePropDescr.xml");
                    _objectsDescriptionExternal = (List<BacnetObjectDescription>)xs.Deserialize(sr);
                }
                catch { }

            }

            valueList = null;

            IList<BacnetPropertyValue> values = new List<BacnetPropertyValue>();

            var oldRetries = comm.Retries;
            comm.Retries = 1;       //we don't want to spend too much time on non existing properties
            try
            {
                // PROP_LIST was added as an addendum to 135-2010
                // Test to see if it is supported, otherwise fall back to the the predefined delault property list.
                var objectDidSupplyPropertyList = ReadProperty(comm, adr, objectId, BacnetPropertyIds.PROP_PROPERTY_LIST, ref values);

                //Used the supplied list of supported Properties, otherwise fall back to using the list of default properties.
                if (objectDidSupplyPropertyList)
                {
                    var propList = values.Last();
                    foreach (var enumeratedValue in propList.value)
                    {
                        BacnetPropertyIds bpi = (BacnetPropertyIds)(uint)enumeratedValue.Value;
                        // read all specified properties given by the xml file
                        ReadProperty(comm, adr, objectId, bpi, ref values);
                    }
                }
                else
                {
                    // Three mandatory common properties to all objects : PROP_OBJECT_IDENTIFIER,PROP_OBJECT_TYPE, PROP_OBJECT_NAME

                    // ReadProperty(comm, adr, object_id, BacnetPropertyIds.PROP_OBJECT_IDENTIFIER, ref values)
                    // No need to query it, known value
                    var newEntry = new BacnetPropertyValue();
                    newEntry.property = new BacnetPropertyReference((uint)BacnetPropertyIds.PROP_OBJECT_IDENTIFIER, System.IO.BACnet.Serialize.ASN1.BACNET_ARRAY_ALL);
                    newEntry.value = new BacnetValue[] { new BacnetValue(objectId) };
                    values.Add(newEntry);

                    // ReadProperty(comm, adr, object_id, BacnetPropertyIds.PROP_OBJECT_TYPE, ref values);
                    // No need to query it, known value
                    newEntry = new BacnetPropertyValue();
                    newEntry.property = new BacnetPropertyReference((uint)BacnetPropertyIds.PROP_OBJECT_TYPE, System.IO.BACnet.Serialize.ASN1.BACNET_ARRAY_ALL);
                    newEntry.value = new BacnetValue[] { new BacnetValue(BacnetApplicationTags.BACNET_APPLICATION_TAG_ENUMERATED, (uint)objectId.type) };
                    values.Add(newEntry);

                    // We do not know the value here
                    ReadProperty(comm, adr, objectId, BacnetPropertyIds.PROP_OBJECT_NAME, ref values);

                    // for all other properties, the list is comming from the internal or external XML file

                    var objDescr = new BacnetObjectDescription(); ;

                    var idx = -1;
                    // try to find the Object description from the optional external xml file
                    if (_objectsDescriptionExternal != null)
                        idx = _objectsDescriptionExternal.FindIndex(o => o.typeId == objectId.type);

                    if (idx != -1)
                        objDescr = _objectsDescriptionExternal[idx];
                    else
                    {
                        // try to find from the embedded resoruce
                        idx = _objectsDescriptionDefault.FindIndex(o => o.typeId == objectId.type);
                        if (idx != -1)
                            objDescr = _objectsDescriptionDefault[idx];
                    }

                    if (idx != -1)
                        foreach (var bpi in objDescr.propsId)
                            // read all specified properties given by the xml file
                            ReadProperty(comm, adr, objectId, bpi, ref values);
                }
            }
            catch { }

            comm.Retries = oldRetries;
            valueList = new BacnetReadAccessResult[] { new BacnetReadAccessResult(objectId, values) };
            return true;
        }

        private void UpdateGrid(TreeNode selectedNode)
        {
            Cursor = Cursors.WaitCursor;
            try
            {
                //fetch end point
                if (m_DeviceTree.SelectedNode == null) return;
                else if (m_DeviceTree.SelectedNode.Tag == null) return;
                else if (!(m_DeviceTree.SelectedNode.Tag is KeyValuePair<BacnetAddress, uint>)) return;
                var entry = (KeyValuePair<BacnetAddress, uint>)m_DeviceTree.SelectedNode.Tag;
                var adr = entry.Key;
                BacnetClient comm;

                if (m_DeviceTree.SelectedNode.Parent.Tag is BacnetClient)
                    comm = (BacnetClient)m_DeviceTree.SelectedNode.Parent.Tag;
                else
                    comm = (BacnetClient)m_DeviceTree.SelectedNode.Parent.Parent.Tag;  // routed node

                if (selectedNode.Tag is BacnetObject)
                {
                    m_DataGrid.SelectedObject = null;   //clear

                    var objectId = (BacnetObject)selectedNode.Tag;
                    var properties = new BacnetPropertyReference[] { new BacnetPropertyReference((uint)BacnetPropertyIds.PROP_ALL, System.IO.BACnet.Serialize.ASN1.BACNET_ARRAY_ALL) };
                    IList<BacnetReadAccessResult> multiValueList;
                    try
                    {
                        //fetch properties. This might not be supported (ReadMultiple) or the response might be too long.
                        if (!comm.ReadPropertyMultipleRequest(adr, objectId, properties, out multiValueList))
                        {
                            Trace.TraceWarning("Couldn't perform ReadPropertyMultiple ... Trying ReadProperty instead");
                            if (!ReadAllPropertiesBySingle(comm, adr, objectId, out multiValueList))
                            {
                                MessageBox.Show(this, "Couldn't fetch properties", "Communication Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }
                        }
                    }
                    catch (Exception)
                    {
                        Trace.TraceWarning("Couldn't perform ReadPropertyMultiple ... Trying ReadProperty instead");
                        Application.DoEvents();
                        try
                        {
                            //fetch properties with single calls
                            if (!ReadAllPropertiesBySingle(comm, adr, objectId, out multiValueList))
                            {
                                MessageBox.Show(this, "Couldn't fetch properties", "Communication Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(this, "Error during read: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }

                    //update grid
                    var bag = new Utilities.DynamicPropertyGridContainer();
                    foreach (BacnetPropertyValue pValue in multiValueList[0].values)
                    {
                        object value = null;
                        BacnetValue[] bValues = null;
                        if (pValue.value != null)
                        {

                            bValues = new BacnetValue[pValue.value.Count];

                            pValue.value.CopyTo(bValues, 0);
                            if (bValues.Length > 1)
                            {
                                var arr = new object[bValues.Length];
                                for (var j = 0; j < arr.Length; j++)
                                    arr[j] = bValues[j].Value;
                                value = arr;
                            }
                            else if (bValues.Length == 1)
                                value = bValues[0].Value;
                        }
                        else
                            bValues = new BacnetValue[0];

                        // Modif FC
                        switch ((BacnetPropertyIds)pValue.property.propertyIdentifier)
                        {
                            // PROP_RELINQUISH_DEFAULT can be write to null value
                            case BacnetPropertyIds.PROP_PRESENT_VALUE:
                                // change to the related nullable type
                                Type t = null;
                                try
                                {
                                    t = value.GetType();
                                    t = Type.GetType("System.Nullable`1[" + value.GetType().FullName + "]");
                                }
                                catch { }
                                bag.Add(new Utilities.CustomProperty(GetNiceName((BacnetPropertyIds)pValue.property.propertyIdentifier), value, t != null ? t : typeof(string), false, "", bValues.Length > 0 ? bValues[0].Tag : (BacnetApplicationTags?)null, null, pValue.property));
                                break;

                            default:
                                bag.Add(new Utilities.CustomProperty(GetNiceName((BacnetPropertyIds)pValue.property.propertyIdentifier), value, value != null ? value.GetType() : typeof(string), false, "", bValues.Length > 0 ? bValues[0].Tag : (BacnetApplicationTags?)null, null, pValue.property));
                                break;
                        }

                        // The Prop Name replace the PropId into the Treenode 
                        if (pValue.property.propertyIdentifier == (byte)BacnetPropertyIds.PROP_OBJECT_NAME)
                        {

                            ChangeTreeNodePropertyName(selectedNode, value.ToString());// Update the object name if needed

                            lock (DevicesObjectsName)
                            {
                                var t = new Tuple<string, BacnetObject>(adr.FullHashString(), objectId);
                                DevicesObjectsName.Remove(t);
                                DevicesObjectsName.Add(t, value.ToString());
                            }
                        }
                    }
                    m_DataGrid.SelectedObject = bag;
                }
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }
        // Fixed a small problem when a right click is down in a Treeview
        private void TreeView_MouseDown(object sender, MouseEventArgs e)
        {
            //if (e.Button != MouseButtons.Right)
            //    return;
            // Store the selected node (can deselect a node).
            //(sender as TreeView).SelectedNode = (sender as TreeView).GetNodeAt(e.X, e.Y);
        }

        private void m_AddressSpaceTree_AfterSelect(object sender, TreeViewEventArgs e)
        {
            UpdateGrid(e.Node);
            BacnetClient cl; BacnetAddress ba; BacnetObject objId;

            // Hide all elements in the toolstrip menu
            foreach (var its in m_AddressSpaceMenuStrip.Items)
                (its as ToolStripMenuItem).Visible = false;
            // Set Subscribe always visible
            m_AddressSpaceMenuStrip.Items[0].Visible = true;

            // Get the node type
            GetObjectLink(out cl, out ba, out objId, BacnetObjectTypes.MAX_BACNET_OBJECT_TYPE);
            // Set visible some elements depending of the object type
            switch (objId.type)
            {
                case BacnetObjectTypes.OBJECT_FILE:
                    m_AddressSpaceMenuStrip.Items[1].Visible = true;
                    m_AddressSpaceMenuStrip.Items[2].Visible = true;
                    break;

                case BacnetObjectTypes.OBJECT_TRENDLOG:
                case BacnetObjectTypes.OBJECT_TREND_LOG_MULTIPLE:
                    m_AddressSpaceMenuStrip.Items[3].Visible = true;
                    break;

                case BacnetObjectTypes.OBJECT_SCHEDULE:
                    m_AddressSpaceMenuStrip.Items[4].Visible = true;
                    break;

                case BacnetObjectTypes.OBJECT_NOTIFICATION_CLASS:
                    m_AddressSpaceMenuStrip.Items[5].Visible = true;
                    break;

                case BacnetObjectTypes.OBJECT_CALENDAR:
                    m_AddressSpaceMenuStrip.Items[6].Visible = true;
                    break;
            }

            // Allows delete menu 
            if (objId.type != BacnetObjectTypes.OBJECT_DEVICE)
                m_AddressSpaceMenuStrip.Items[7].Visible = true;
        }

        private void m_DataGrid_PropertyValueChanged(object s, PropertyValueChangedEventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            try
            {
                //fetch end point
                if (m_DeviceTree.SelectedNode == null) return;
                else if (m_DeviceTree.SelectedNode.Tag == null) return;
                else if (!(m_DeviceTree.SelectedNode.Tag is KeyValuePair<BacnetAddress, uint>)) return;
                var entry = (KeyValuePair<BacnetAddress, uint>)m_DeviceTree.SelectedNode.Tag;
                var adr = entry.Key;

                BacnetClient comm;

                if (m_DeviceTree.SelectedNode.Parent.Tag is BacnetClient)
                    comm = (BacnetClient)m_DeviceTree.SelectedNode.Parent.Tag;
                else
                    comm = (BacnetClient)m_DeviceTree.SelectedNode.Parent.Parent.Tag; // a node under a router

                //fetch object_id
                if (m_AddressSpaceTree.SelectedNode == null) return;
                else if (m_AddressSpaceTree.SelectedNode.Tag == null) return;
                else if (!(m_AddressSpaceTree.SelectedNode.Tag is BacnetObject)) return;
                var objectId = (BacnetObject)m_AddressSpaceTree.SelectedNode.Tag;

                Utilities.CustomPropertyDescriptor c = null;
                var gridItem = e.ChangedItem;
                // Go up to the Property (could be a sub-element)

                do
                {
                    if (gridItem.PropertyDescriptor is Utilities.CustomPropertyDescriptor)
                        c = (Utilities.CustomPropertyDescriptor)gridItem.PropertyDescriptor;
                    else
                        gridItem = gridItem.Parent;

                } while ((c == null) && (gridItem != null));

                if (c == null) return; // never occur normally

                //fetch property
                var property = (BacnetPropertyReference)c.CustomProperty.Tag;
                //new value
                var newValue = gridItem.Value;

                //convert to bacnet
                BacnetValue[] bValue = null;
                try
                {
                    if (newValue != null && newValue.GetType().IsArray)
                    {
                        var arr = (Array)newValue;
                        bValue = new BacnetValue[arr.Length];
                        for (var i = 0; i < arr.Length; i++)
                            bValue[i] = new BacnetValue(arr.GetValue(i));
                    }
                    else
                    {
                        {
                            // Modif FC
                            bValue = new BacnetValue[1];
                            if ((BacnetApplicationTags)c.CustomProperty.bacnetApplicationTags != BacnetApplicationTags.BACNET_APPLICATION_TAG_NULL)
                            {
                                bValue[0] = new BacnetValue((BacnetApplicationTags)c.CustomProperty.bacnetApplicationTags, newValue);
                            }
                            else
                            {
                                object o = null;
                                var t = new TypeConverter();
                                // try to convert to the simplest type
                                string[] typeList = { "Boolean", "UInt32", "Int32", "Single", "Double" };

                                foreach (var typename in typeList)
                                {
                                    try
                                    {
                                        o = Convert.ChangeType(newValue, Type.GetType("System." + typename));
                                        break;
                                    }
                                    catch { }
                                }

                                if (o == null)
                                    bValue[0] = new BacnetValue(newValue);
                                else
                                    bValue[0] = new BacnetValue(o);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Couldn't convert property: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                //write
                try
                {
                    comm.WritePriority = (uint)Properties.Settings.Default.DefaultWritePriority;
                    if (!comm.WritePropertyRequest(adr, objectId, (BacnetPropertyIds)property.propertyIdentifier, bValue))
                    {
                        MessageBox.Show(this, "Couldn't write property", "Communication Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Error during write: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                //reload
                UpdateGrid(m_AddressSpaceTree.SelectedNode);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        public bool GetObjectLink(out BacnetClient comm, out BacnetAddress adr, out BacnetObject objectId, BacnetObjectTypes expectedType)
        {

            comm = null;
            adr = new BacnetAddress(BacnetAddressTypes.None, 0, null);
            objectId = new BacnetObject();

            try
            {
                if (m_DeviceTree.SelectedNode == null) return false;
                else if (m_DeviceTree.SelectedNode.Tag == null) return false;
                else if (!(m_DeviceTree.SelectedNode.Tag is KeyValuePair<BacnetAddress, uint>)) return false;
                var entry = (KeyValuePair<BacnetAddress, uint>)m_DeviceTree.SelectedNode.Tag;
                adr = entry.Key;
                if (m_DeviceTree.SelectedNode.Parent.Tag is BacnetClient)
                    comm = (BacnetClient)m_DeviceTree.SelectedNode.Parent.Tag;
                else  // a routed node
                    comm = (BacnetClient)m_DeviceTree.SelectedNode.Parent.Parent.Tag;
            }
            catch
            {
                if (expectedType != BacnetObjectTypes.MAX_BACNET_OBJECT_TYPE)
                    MessageBox.Show(this, "This is not a valid node", "Wrong node", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            //fetch object_id
            if (
                m_AddressSpaceTree.SelectedNode == null ||
                !(m_AddressSpaceTree.SelectedNode.Tag is BacnetObject) ||
                ((BacnetObject)m_AddressSpaceTree.SelectedNode.Tag).type != expectedType)
            {
                var s = expectedType.ToString().Substring(7).ToLower();
                if (expectedType != BacnetObjectTypes.MAX_BACNET_OBJECT_TYPE)
                {
                    MessageBox.Show(this, "The marked object is not a " + s, s, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return false;
                }
            }

            if (m_AddressSpaceTree.SelectedNode != null)
            {
                objectId = (BacnetObject)m_AddressSpaceTree.SelectedNode.Tag;
                return true;
            }

            return false;
        }

        private void downloadFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            try
            {
                //fetch end point
                BacnetClient client = null;
                BacnetAddress address;
                BacnetObject objectId;
                if (GetObjectLink(out client, out address, out objectId, BacnetObjectTypes.OBJECT_FILE) == false) return;

                //where to store file?
                var dlg = new SaveFileDialog();
                dlg.FileName = Properties.Settings.Default.GUI_LastFilename;
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                var filename = dlg.FileName;
                Properties.Settings.Default.GUI_LastFilename = filename;

                //get file size
                var filesize = FileTransfers.ReadFileSize(client, address, objectId);
                if (filesize < 0)
                {
                    MessageBox.Show(this, "Couldn't read file size", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                //display progress
                var progress = new ProgressDialog();
                progress.Text = "Downloading file ...";
                progress.Label = "0 of " + (filesize / 1024) + " kb ... (0.0 kb/s)";
                progress.Maximum = filesize;
                progress.Show(this);

                var start = DateTime.Now;
                double kbPerSec = 0;
                var transfer = new FileTransfers();
                EventHandler cancelHandler = (s, a) => { transfer.Cancel = true; };
                progress.Cancel += cancelHandler;
                Action<int> updateProgress = (position) =>
                {
                    kbPerSec = (position / 1024) / (DateTime.Now - start).TotalSeconds;
                    progress.Value = position;
                    progress.Label = string.Format((position / 1024) + " of " + (filesize / 1024) + " kb ... ({0:F1} kb/s)", kbPerSec);
                };
                Application.DoEvents();
                try
                {
                    if (Properties.Settings.Default.DefaultDownloadSpeed == 2)
                        transfer.DownloadFileBySegmentation(client, address, objectId, filename, updateProgress);
                    else if (Properties.Settings.Default.DefaultDownloadSpeed == 1)
                        transfer.DownloadFileByAsync(client, address, objectId, filename, updateProgress);
                    else
                        transfer.DownloadFileByBlocking(client, address, objectId, filename, updateProgress);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Error during download file: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                finally
                {
                    progress.Hide();
                }
            }
            finally
            {
                Cursor = Cursors.Default;
            }

            //information
            try
            {
                MessageBox.Show(this, "Done", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch
            {
            }
        }

        private void uploadFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            try
            {
                //fetch end point
                BacnetClient client = null;
                BacnetAddress address;
                BacnetObject objectId;
                if (GetObjectLink(out client, out address, out objectId, BacnetObjectTypes.OBJECT_FILE) == false) return;

                //which file to upload?
                var dlg = new OpenFileDialog {FileName = Properties.Settings.Default.GUI_LastFilename};
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                var filename = dlg.FileName;
                Properties.Settings.Default.GUI_LastFilename = filename;

                //display progress
                var fileSize = (int)(new FileInfo(filename)).Length;
                var progress = new ProgressDialog
                {
                    Text = "Uploading file ...",
                    Label = "0 of " + (fileSize / 1024) + " kb ... (0.0 kb/s)",
                    Maximum = fileSize
                };
                progress.Show(this);

                var transfer = new FileTransfers();
                var start = DateTime.Now;
                double kbPerSec = 0;
                EventHandler cancelHandler = (s, a) => { transfer.Cancel = true; };
                progress.Cancel += cancelHandler;
                Action<int> updateProgress = (position) =>
                {
                    kbPerSec = (position / 1024) / (DateTime.Now - start).TotalSeconds;
                    progress.Value = position;
                    progress.Label = string.Format((position / 1024) + " of " + (fileSize / 1024) + " kb ... ({0:F1} kb/s)", kbPerSec);
                };
                try
                {
                    transfer.UploadFileByBlocking(client, address, objectId, filename, updateProgress);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "Error during upload file: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                finally
                {
                    progress.Hide();
                }
            }
            finally
            {
                Cursor = Cursors.Default;
            }

            //information
            MessageBox.Show(this, "Done", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // FC
        private void showTrendLogToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                //fetch end point
                BacnetClient client;
                BacnetAddress address;
                BacnetObject objectId;

                if (GetObjectLink(out client, out address, out objectId, BacnetObjectTypes.OBJECT_TRENDLOG) == false)
                    if (GetObjectLink(out client, out address, out objectId, BacnetObjectTypes.OBJECT_TREND_LOG_MULTIPLE) == false) return;

                new TrendLogDisplay(client, address, objectId).ShowDialog();

            }
            catch (Exception ex)
            {
                Trace.TraceError("Error loading TrendLog : " + ex.Message);
            }
        }
        // FC
        private void showScheduleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                //fetch end point
                BacnetClient client;
                BacnetAddress address;
                BacnetObject objectId;

                if (GetObjectLink(out client, out address, out objectId, BacnetObjectTypes.OBJECT_SCHEDULE) == false) return;

                new ScheduleDisplay(m_AddressSpaceTree.ImageList, client, address, objectId).ShowDialog();

            }
            catch (Exception ex) { Trace.TraceError("Error loading Schedule : " + ex.Message); }
        }

        private void deleteObjectToolStripMenuItem_Click(object sender, EventArgs e)
        {

            try
            {
                //fetch end point
                BacnetClient client;
                BacnetAddress address;
                BacnetObject objectId;

                GetObjectLink(out client, out address, out objectId, BacnetObjectTypes.MAX_BACNET_OBJECT_TYPE);

                if (MessageBox.Show("Are you sure you want to delete this object ?", objectId.ToString(), MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    client.DeleteObjectRequest(address, objectId);
                    m_DeviceTree_AfterSelect(null, new TreeViewEventArgs(m_DeviceTree.SelectedNode));
                }

            }
            catch (Exception ex)
            {
                Trace.TraceError("Error : " + ex.Message);
                MessageBox.Show("Fail to Delete Object", "DeleteObject", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void showCalendarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                //fetch end point
                BacnetClient client;
                BacnetAddress address;
                BacnetObject objectId;

                if (GetObjectLink(out client, out address, out objectId, BacnetObjectTypes.OBJECT_CALENDAR) == false) return;

                new CalendarEditor(client, address, objectId).ShowDialog();

            }
            catch (Exception ex) { Trace.TraceError("Error loading Calendar : " + ex.Message); }
        }

        //FC
        private void showNotificationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                //fetch end point
                BacnetClient client;
                BacnetAddress address;
                BacnetObject objectId;

                if (GetObjectLink(out client, out address, out objectId, BacnetObjectTypes.OBJECT_NOTIFICATION_CLASS) == false) return;

                new NotificationEditor(client, address, objectId).ShowDialog();

            }
            catch (Exception ex) { Trace.TraceError("Error loading Notification : " + ex.Message); }
        }


        private void m_AddressSpaceTree_ItemDrag(object sender, ItemDragEventArgs e)
        {
            m_AddressSpaceTree.DoDragDrop(e.Item, DragDropEffects.Move);
        }

        private void m_SubscriptionView_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private string GetObjectName(BacnetClient comm, BacnetAddress adr, BacnetObject objectId)
        {
            Cursor = Cursors.WaitCursor;
            try
            {
                IList<BacnetValue> value;
                if (!comm.ReadPropertyRequest(adr, objectId, BacnetPropertyIds.PROP_OBJECT_NAME, out value))
                    return "[Timed out]";
                if (value == null || value.Count == 0)
                    return "";
                else
                    return value[0].Value.ToString();
            }
            catch (Exception ex)
            {
                return "[Error: " + ex.Message + " ]";
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private class Subscription
        {
            public BacnetClient Client;
            public BacnetAddress Address;
            public BacnetObject DeviceId, ObjectId;
            public string SubscriptionKey;
            public uint SubscriptionId;
            public bool IsActiveSubscription = true; // false if subscription is refused

            public Subscription(BacnetClient comm, BacnetAddress adr, BacnetObject deviceId, BacnetObject objectId, string subKey, uint subscribeId)
            {
                Client = comm;
                Address = adr;
                DeviceId = deviceId;
                ObjectId = objectId;
                SubscriptionKey = subKey;
                SubscriptionId = subscribeId;
            }
        }

        private bool CreateSubscription(BacnetClient comm, BacnetAddress adr, uint deviceId, BacnetObject objectId, bool withGraph)
        {
            Cursor = Cursors.WaitCursor;
            try
            {
                //fetch device_id if needed
                if (deviceId >= System.IO.BACnet.Serialize.ASN1.BACNET_MAX_INSTANCE)
                {
                    deviceId = FetchDeviceId(comm, adr);
                }

                _mNextSubscriptionId++;
                var subKey = adr + ":" + deviceId + ":" + _mNextSubscriptionId;
                var sub = new Subscription(comm, adr, new BacnetObject(BacnetObjectTypes.OBJECT_DEVICE, deviceId), objectId, subKey, _mNextSubscriptionId);

                //add to list
                var itm = m_SubscriptionView.Items.Add(adr + " - " + deviceId);
                itm.SubItems.Add(objectId.ToString());
                itm.SubItems.Add(GetObjectName(comm, adr, objectId));   //name
                itm.SubItems.Add("");   //value
                itm.SubItems.Add("");   //time
                itm.SubItems.Add("Not started");   //status
                if (Properties.Settings.Default.ShowDescriptionWhenUsefull)
                {
                    IList<BacnetValue> values;
                    if (comm.ReadPropertyRequest(adr, objectId, BacnetPropertyIds.PROP_DESCRIPTION, out values))
                    {
                        itm.SubItems.Add(values[0].Value.ToString());   // Description
                    }
                }
                else
                    itm.SubItems.Add(""); // Description

                itm.SubItems.Add("");   // Graph Line Color
                itm.Tag = sub;

                lock (_mSubscriptionList)
                {
                    _mSubscriptionList.Add(subKey, itm);
                    if (withGraph)
                    {
                        var points = new RollingPointPairList(1000);
                        _mSubscriptionPoints.Add(subKey, points);
                        var color = _graphColor[_pane.CurveList.Count % _graphColor.Length];
                        var l = _pane.AddCurve("", points, color, Properties.Settings.Default.GraphDotStyle);
                        l.Tag = points; // store the link To be able to remove the LineItem
                        itm.SubItems[7].BackColor = color;
                        itm.UseItemStyleForSubItems = false;
                        CovGraph.Invalidate();
                    }
                }

                //add to device

                var subscribeOk = false;

                try
                {
                    subscribeOk = comm.SubscribeCOVRequest(adr, objectId, _mNextSubscriptionId, false, Properties.Settings.Default.Subscriptions_IssueConfirmedNotifies, Properties.Settings.Default.Subscriptions_Lifetime);
                }
                catch { }

                if (subscribeOk == false) // echec : launch period acquisiton in the ThreadPool
                {
                    sub.IsActiveSubscription = false;
                    var qst = new GenericInputBox<NumericUpDown>("Error during subscribe", "Polling period replacement (s)",
                              (o) =>
                              {
                                  o.Minimum = 1; o.Maximum = 120; o.Value = Properties.Settings.Default.Subscriptions_ReplacementPollingPeriod;
                              });

                    var rep = qst.ShowDialog();
                    if (rep == DialogResult.OK)
                    {
                        var period = (int)qst.genericInput.Value;
                        Properties.Settings.Default.Subscriptions_ReplacementPollingPeriod = (uint)period;
                        ThreadPool.QueueUserWorkItem(a => ReadPropertyPollingReplacementToCov(sub, period));
                    }

                    return false; // COV is not done
                }
            }
            catch
            {
                return false;
            }
            finally
            {
                Cursor = Cursors.Default;
            }

            return true;
        }

        // COV echec, PROP_PRESENT_VALUE read replacement method
        // x seconds polling period
        private void ReadPropertyPollingReplacementToCov(Subscription sub, int period)
        {
            for (; ; )
            {
                IList<BacnetPropertyValue> values = new List<BacnetPropertyValue>();
                if (ReadProperty(sub.Client, sub.Address, sub.ObjectId, BacnetPropertyIds.PROP_PRESENT_VALUE, ref values) == false)
                    return; // maybe here we could not go away 

                lock (_mSubscriptionList)
                    if (_mSubscriptionList.ContainsKey(sub.SubscriptionKey))
                        // COVNotification replacement
                        OnCOVNotification(sub.Client, sub.Address, 0, sub.SubscriptionId, sub.DeviceId, sub.ObjectId, 0, false, values, BacnetMaxSegments.MAX_SEG0);
                    else
                        return;

                Thread.Sleep(Math.Max(1, period) * 1000);
            }
        }

        private void m_SubscriptionView_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent("CodersLab.Windows.Controls.NodesCollection", false))
            {
                //fetch end point
                if (m_DeviceTree.SelectedNode == null) return;
                else if (m_DeviceTree.SelectedNode.Tag == null) return;
                else if (!(m_DeviceTree.SelectedNode.Tag is KeyValuePair<BacnetAddress, uint>)) return;
                var entry = (KeyValuePair<BacnetAddress, uint>)m_DeviceTree.SelectedNode.Tag;
                var adr = entry.Key;

                BacnetClient comm;
                if (m_DeviceTree.SelectedNode.Parent.Tag is BacnetClient)
                    comm = (BacnetClient)m_DeviceTree.SelectedNode.Parent.Tag;
                else  // a routed device
                    comm = (BacnetClient)m_DeviceTree.SelectedNode.Parent.Parent.Tag;

                //fetch object_id
                var nodes = (CodersLab.Windows.Controls.NodesCollection)e.Data.GetData("CodersLab.Windows.Controls.NodesCollection");
                //node[0]

                // Nodes are in a non controllable order, so puts the objectIds in order
                var bacnetObjects = new List<BacnetObject>();
                for (var i = 0; i < nodes.Count; i++)
                {
                    if ((nodes[i].Tag != null) && (nodes[i].Tag is BacnetObject))
                        bacnetObjects.Add((BacnetObject)nodes[i].Tag);
                }

                bacnetObjects.Sort();

                for (var i = 0; i < bacnetObjects.Count; i++)
                {
                    if (CreateSubscription(comm, adr, entry.Value, bacnetObjects[i], sender == CovGraph) == false)
                        break;
                }
            }
        }

        private void MainDialog_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                //commit setup
                Properties.Settings.Default.GUI_SplitterButtom = m_SplitContainerButtom.SplitterDistance;
                Properties.Settings.Default.GUI_SplitterLeft = m_SplitContainerLeft.SplitterDistance;
                Properties.Settings.Default.GUI_SplitterRight = m_SplitContainerRight.SplitterDistance;
                Properties.Settings.Default.GUI_FormSize = Size;
                Properties.Settings.Default.GUI_FormState = WindowState.ToString();

                //save
                Properties.Settings.Default.Save();

                // save object name<->id file
                Stream stream = File.Open(Properties.Settings.Default.ObjectNameFile, FileMode.Create);
                var bf = new BinaryFormatter();
                bf.Serialize(stream, DevicesObjectsName);
                stream.Close();

            }
            catch
            {
                //ignore
            }
        }

        private void m_SubscriptionView_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Delete) return;

            if (m_SubscriptionView.SelectedItems.Count >= 1)
            {
                foreach (ListViewItem itm in m_SubscriptionView.SelectedItems)
                {
                    //ListViewItem itm = m_SubscriptionView.SelectedItems[0];
                    if (itm.Tag is Subscription)    // It's a subscription or not (Event/Alarm)
                    {
                        var sub = (Subscription)itm.Tag;
                        if (_mSubscriptionList.ContainsKey(sub.SubscriptionKey))
                        {
                            //remove from device
                            try
                            {
                                if (sub.IsActiveSubscription)
                                    if (!sub.Client.SubscribeCOVRequest(sub.Address, sub.ObjectId, sub.SubscriptionId, true, false, 0))
                                    {
                                        MessageBox.Show(this, "Couldn't unsubscribe", "Communication Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                        return;
                                    }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(this, "Couldn't delete subscription: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                        }
                        //remove from interface
                        m_SubscriptionView.Items.Remove(itm);
                        lock (_mSubscriptionList)
                        {
                            _mSubscriptionList.Remove(sub.SubscriptionKey);
                            try
                            {
                                var points = _mSubscriptionPoints[sub.SubscriptionKey];
                                foreach (LineItem l in _pane.CurveList)
                                    if (l.Tag == points)
                                    {
                                        _pane.CurveList.Remove(l);
                                        break;
                                    }

                                _mSubscriptionPoints.Remove(sub.SubscriptionKey);
                            }
                            catch { }
                        }

                        CovGraph.AxisChange();
                        CovGraph.Invalidate();
                        m_SubscriptionView.Items.Remove(itm);
                    }

                }
            }
        }

        private void sendWhoIsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                var comm = (BacnetClient)m_DeviceTree.SelectedNode.Tag;
                comm.WhoIs();
            }
            catch
            {
                MessageBox.Show(this, "Please select a \"transport\" node first", "Wrong node", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void AddRemoteIpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //fetch end point
            BacnetClient client = null;
            try
            {
                if (m_DeviceTree.SelectedNode == null) return;
                else if (m_DeviceTree.SelectedNode.Tag == null) return;
                else if (!(m_DeviceTree.SelectedNode.Tag is BacnetClient)) return;
                client = (BacnetClient)m_DeviceTree.SelectedNode.Tag;

                if (client.Transport is BacnetIpUdpProtocolTransport) // only IPv4 today, v6 maybe a day
                {

                    var input =
                        new GenericInputBox<TextBox>("Ipv4/Udp Bacnet Node", "DeviceId - xx.xx.xx.xx:47808",
                          (o) =>
                          {
                              // adjustment to the generic control
                          }, 1, true, "Unknown device Id can be replaced by 4194303 or ?");
                    var res = input.ShowDialog();

                    if (res == DialogResult.OK)
                    {
                        var entry = input.genericInput.Text.Split('-');
                        if (entry[0][0] == '?') entry[0] = "4194303";
                        OnIam(client, new BacnetAddress(BacnetAddressTypes.IP, entry[1].Trim()), Convert.ToUInt32(entry[0]), 0, BacnetSegmentations.SEGMENTATION_NONE, 0);
                    }
                }
                else
                {
                    MessageBox.Show(this, "Please select an \"IPv4 transport\" node first", "Wrong node", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch
            {
                MessageBox.Show(this, "Invalid parameter", "Wrong node or IP @", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void helpToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            var readmePath = Path.GetDirectoryName(Application.ExecutablePath) + "/README.txt";
            Process.Start(readmePath);
        }

        private void settingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var dlg = new SettingsDialog {SelectedObject = Properties.Settings.Default};
            dlg.ShowDialog(this);
        }

        /// <summary>
        /// This will download all values from a given device and store it in a xml format, fit for the DemoServer
        /// This can be a good way to test serializing
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void exportDeviceDBToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //fetch end point
            BacnetClient client;
            BacnetAddress address;
            uint deviceId;

            FetchEndPoint(out client, out address, out deviceId);

            if (client == null)
            {
                MessageBox.Show(this, "Please select a device node", "Wrong node", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }


            //select file to store
            var dlg = new SaveFileDialog {Filter = "xml|*.xml"};
            if (dlg.ShowDialog(this) != DialogResult.OK) return;

            Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            var removeObject = false;

            try
            {
                //get all objects
                var storage = new DeviceStorage();
                IList<BacnetValue> valueList;
                client.ReadPropertyRequest(address, new BacnetObject(BacnetObjectTypes.OBJECT_DEVICE, deviceId), BacnetPropertyIds.PROP_OBJECT_LIST, out valueList);
                var objectList = new LinkedList<BacnetObject>();
                foreach (var value in valueList)
                {
                    if (Enum.IsDefined(typeof(BacnetObjectTypes), ((BacnetObject)value.Value).Type))
                        objectList.AddLast((BacnetObject)value.Value);
                    else
                        removeObject = true;
                }

                foreach (var objectId in objectList)
                {
                    //read all properties
                    IList<BacnetReadAccessResult> multiValueList;
                    var properties = new BacnetPropertyReference[] { new BacnetPropertyReference((uint)BacnetPropertyIds.PROP_ALL, System.IO.BACnet.Serialize.ASN1.BACNET_ARRAY_ALL) };
                    client.ReadPropertyMultipleRequest(address, objectId, properties, out multiValueList);

                    //store
                    foreach (BacnetPropertyValue value in multiValueList[0].values)
                    {
                        try
                        {
                            storage.WriteProperty(objectId, (BacnetPropertyIds)value.property.propertyIdentifier, value.property.propertyArrayIndex, value.value, true);
                        }
                        catch { }
                    }
                }

                //save to disk
                storage.Save(dlg.FileName);

                //display
                MessageBox.Show(this, "Done", "Export done", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error during export: " + ex.Message, "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
                if (removeObject == true)
                    Trace.TraceWarning("All proprietary Objects removed from export");
            }
        }

        private uint FetchDeviceId(BacnetClient comm, BacnetAddress adr)
        {
            IList<BacnetValue> value;
            if (comm.ReadPropertyRequest(adr, new BacnetObject(BacnetObjectTypes.OBJECT_DEVICE, System.IO.BACnet.Serialize.ASN1.BACNET_MAX_INSTANCE), BacnetPropertyIds.PROP_OBJECT_IDENTIFIER, out value))
            {
                if (value != null && value.Count > 0 && value[0].Value is BacnetObject)
                {
                    var objectId = (BacnetObject)value[0].Value;
                    return objectId.instanceId;
                }
                else
                    return 0xFFFFFFFF;
            }
            else
                return 0xFFFFFFFF;
        }

        private void subscribeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //fetch end point
            if (m_DeviceTree.SelectedNode == null) return;
            else if (m_DeviceTree.SelectedNode.Tag == null) return;
            else if (!(m_DeviceTree.SelectedNode.Tag is KeyValuePair<BacnetAddress, uint>)) return;
            var entry = (KeyValuePair<BacnetAddress, uint>)m_DeviceTree.SelectedNode.Tag;
            var adr = entry.Key;
            BacnetClient comm;
            if (m_DeviceTree.SelectedNode.Parent.Tag is BacnetClient)
                comm = (BacnetClient)m_DeviceTree.SelectedNode.Parent.Tag;
            else
                comm = (BacnetClient)m_DeviceTree.SelectedNode.Parent.Parent.Tag; // When device is under a Router
            var deviceId = entry.Value;

            //test object_id with the last selected node
            if (
                m_AddressSpaceTree.SelectedNode == null ||
                !(m_AddressSpaceTree.SelectedNode.Tag is BacnetObject))
            {
                MessageBox.Show(this, "The marked object is not an object", "Not an object", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            // advise all selected nodes, stop at the first COV reject (even if a period polling is done)
            foreach (TreeNode t in m_AddressSpaceTree.SelectedNodes)
            {
                var objectId = (BacnetObject)t.Tag;
                //create 
                if (CreateSubscription(comm, adr, deviceId, objectId, false) == false)
                    return;
            }
        }

        private void m_subscriptionRenewTimer_Tick(object sender, EventArgs e)
        {
            // don't want to lock the list for a while
            // so get element one by one using the indexer            
            int itmCount;
            lock (_mSubscriptionList)
                itmCount = _mSubscriptionList.Count;

            for (var i = 0; i < itmCount; i++)
            {
                ListViewItem itm = null;

                // lock another time the list to get the item by indexer
                try
                {
                    lock (_mSubscriptionList)
                        itm = _mSubscriptionList.Values.ElementAt(i);
                }
                catch { }

                if (itm != null)
                {
                    try
                    {
                        var sub = (Subscription)itm.Tag;

                        if (sub.IsActiveSubscription == false) // not needs to renew, periodic pooling in operation (or nothing) due to COV subscription refused by the remote device
                            return;

                        if (!sub.Client.SubscribeCOVRequest(sub.Address, sub.ObjectId, sub.SubscriptionId, false, Properties.Settings.Default.Subscriptions_IssueConfirmedNotifies, Properties.Settings.Default.Subscriptions_Lifetime))
                        {
                            SetSubscriptionStatus(itm, "Offline");
                            Trace.TraceWarning("Couldn't renew subscription " + sub.SubscriptionId);
                        }
                    }
                    catch (Exception ex)
                    {
                        Trace.TraceError("Exception during renew subscription: " + ex.Message);
                    }
                }
            }
        }

        private void sendWhoIsToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            sendWhoIsToolStripMenuItem_Click(this, null);
        }

        private void exportDeviceDBToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            exportDeviceDBToolStripMenuItem_Click(this, null);
        }

        private void downloadFileToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            downloadFileToolStripMenuItem_Click(this, null);
        }

        private void showTrendLogToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            showTrendLogToolStripMenuItem_Click(null, null);
        }

        private void showScheduleToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            showScheduleToolStripMenuItem_Click(null, null);
        }

        private void showCalendarToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            showCalendarToolStripMenuItem_Click(null, null);
        }

        private void showNotificationToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            showNotificationToolStripMenuItem_Click(null, null);
        }
        private void uploadFileToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            uploadFileToolStripMenuItem_Click(this, null);
        }

        private void subscribeToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            subscribeToolStripMenuItem_Click(this, null);
        }

        private void timeSynchronizeToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            timeSynchronizeToolStripMenuItem_Click(this, null);
        }

        // retrieve the BacnetClient, BacnetAddress, device id of the selected node
        private void FetchEndPoint(out BacnetClient comm, out BacnetAddress adr, out uint deviceId)
        {
            comm = null; adr = null; deviceId = 0;
            try
            {
                if (m_DeviceTree.SelectedNode == null) return;
                else if (m_DeviceTree.SelectedNode.Tag == null) return;
                else if (!(m_DeviceTree.SelectedNode.Tag is KeyValuePair<BacnetAddress, uint>)) return;
                var entry = (KeyValuePair<BacnetAddress, uint>)m_DeviceTree.SelectedNode.Tag;
                adr = entry.Key;
                deviceId = entry.Value;
                if (m_DeviceTree.SelectedNode.Parent.Tag is BacnetClient)
                    comm = (BacnetClient)m_DeviceTree.SelectedNode.Parent.Tag;
                else
                    comm = (BacnetClient)m_DeviceTree.SelectedNode.Parent.Parent.Tag; // When device is under a Router
            }
            catch
            {

            }
        }

        private void timeSynchronizeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //fetch end point
            BacnetClient client;
            BacnetAddress address;
            uint deviceId;

            FetchEndPoint(out client, out address, out deviceId);

            if (client == null)
            {
                MessageBox.Show(this, "Please select a device node", "Wrong node", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            //send
            if (Properties.Settings.Default.TimeSynchronize_UTC)
                client.SynchronizeTime(address, DateTime.Now.ToUniversalTime(), true);
            else
                client.SynchronizeTime(address, DateTime.Now, false);

            //done
            MessageBox.Show(this, "OK", "Time Synchronize", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void communicationControlToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //fetch end point
            BacnetClient client;
            BacnetAddress address;
            uint deviceId;

            FetchEndPoint(out client, out address, out deviceId);

            if (client == null)
            {
                MessageBox.Show(this, "Please select a device node", "Wrong node", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            //Options
            var dlg = new DeviceCommunicationControlDialog();
            if (dlg.ShowDialog(this) != DialogResult.OK) return;

            if (dlg.IsReinitialize)
            {
                //Reinitialize Device
                if (!client.ReinitializeRequest(address, dlg.ReinitializeState, dlg.Password))
                    MessageBox.Show(this, "Couldn't perform device communication control", "Device Communication Control", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                    MessageBox.Show(this, "OK", "Device Communication Control", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                //Device Communication Control
                if (!client.DeviceCommunicationControlRequest(address, dlg.Duration, dlg.DisableCommunication ? (uint)1 : (uint)0, dlg.Password))
                    MessageBox.Show(this, "Couldn't perform device communication control", "Device Communication Control", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                    MessageBox.Show(this, "OK", "Device Communication Control", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void communicationControlToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            communicationControlToolStripMenuItem_Click(this, null);
        }
        // Modify FC
        // base on http://www.big-eu.org/fileadmin/downloads/EDE2_2_Templates.zip
        // This will download all values from a given device and store it in an EDE csv format,
        private void exportDeviceEDEFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //fetch end point
            BacnetClient client;
            BacnetAddress address;
            uint deviceId;

            FetchEndPoint(out client, out address, out deviceId);

            if (client == null)
            {
                MessageBox.Show(this, "Please select a device node", "Wrong node", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            //select file to store
            var dlg = new SaveFileDialog {Filter = "csv|*.csv"};
            if (dlg.ShowDialog(this) != DialogResult.OK) return;

            Cursor = Cursors.WaitCursor;
            Application.DoEvents();
            try
            {
                var sw = new StreamWriter(dlg.FileName);

                sw.WriteLine("# Proposal_Engineering-Data-Exchange - B.I.G.-EU");
                sw.WriteLine("PROJECT_NAME");
                sw.WriteLine("VERSION_OF_REFERENCEFILE");
                sw.WriteLine("TIMESTAMP_OF_LAST_CHANGE;" + DateTime.Now.ToShortDateString());
                sw.WriteLine("AUTHOR_OF_LAST_CHANGE;YABE Yet Another Bacnet Explorer");
                sw.WriteLine("VERSION_OF_LAYOUT;2.2");
                sw.WriteLine("#mandatory;mandatory;mandatory;mandatory;mandatory;optional;optional;optional;optional;optional;optional;optional;optional;optional;optional;optional");
                sw.WriteLine("# keyname;device obj.-instance;object-name;object-type;object-instance;description;present-value-default;min-present-value;max-present-value;settable;supports COV;hi-limit;low-limit;state-text-reference;unit-code;vendor-specific-addres");

                var properties = new BacnetPropertyReference[2]
                                                                {
                                                                    new BacnetPropertyReference((uint)BacnetPropertyIds.PROP_OBJECT_NAME, System.IO.BACnet.Serialize.ASN1.BACNET_ARRAY_ALL),
                                                                    new BacnetPropertyReference((uint)BacnetPropertyIds.PROP_DESCRIPTION, System.IO.BACnet.Serialize.ASN1.BACNET_ARRAY_ALL),
                                                                };

                var readPropertyMultipleSupported = true;

                // Object list is already in the AddressSpaceTree, so no need to query it again
                foreach (TreeNode t in m_AddressSpaceTree.Nodes)
                {
                    var bacnetObject = (BacnetObject)t.Tag;
                    var identifier = "";
                    var description = "";
                    var unitCode = ""; // Not actually in use

                    var propObjectNameOk = false;
                    lock (DevicesObjectsName)
                        propObjectNameOk = DevicesObjectsName.TryGetValue(new Tuple<string, BacnetObject>(address.FullHashString(), bacnetObject), out identifier);

                    if ((readPropertyMultipleSupported) && (!propObjectNameOk))
                    {
                        try
                        {
                            IList<BacnetReadAccessResult> multiValueList;
                            var propToRead = new BacnetReadAccessSpecification[] { new BacnetReadAccessSpecification(bacnetObject, properties) };
                            client.ReadPropertyMultipleRequest(address, propToRead, out multiValueList);
                            var br = multiValueList[0];

                            foreach (BacnetPropertyValue pv in br.values)
                            {
                                if ((BacnetPropertyIds)pv.property.propertyIdentifier == BacnetPropertyIds.PROP_OBJECT_NAME)
                                    identifier = pv.value[0].Value.ToString();
                                if ((BacnetPropertyIds)pv.property.propertyIdentifier == BacnetPropertyIds.PROP_DESCRIPTION)
                                    if (!(pv.value[0].Value is BacnetError))
                                        description = pv.value[0].Value.ToString(); ;
                            }
                        }
                        catch { readPropertyMultipleSupported = false; }
                    }
                    else
                    {
                        IList<BacnetValue> outValue;

                        if (!propObjectNameOk)
                        {
                            client.ReadPropertyRequest(address, bacnetObject, BacnetPropertyIds.PROP_OBJECT_NAME, out outValue);
                            identifier = outValue[0].Value.ToString();
                        }

                        try
                        {
                            client.ReadPropertyRequest(address, bacnetObject, BacnetPropertyIds.PROP_DESCRIPTION, out outValue);
                            if (!(outValue[0].Value is BacnetError))
                                description = outValue[0].Value.ToString();
                        }
                        catch { }
                    }

                    sw.WriteLine(bacnetObject + ";" + deviceId + ";" + identifier + ";" + ((int)bacnetObject.type) + ";" + bacnetObject.instanceId + ";" + description + ";;;;;;;;;" + unitCode);

                    // Update also the Dictionary of known object name and the TreeNode
                    if (t.ToolTipText == "")
                    {
                        lock (DevicesObjectsName)
                            DevicesObjectsName.Add(new Tuple<string, BacnetObject>(address.FullHashString(), bacnetObject), identifier);

                        ChangeTreeNodePropertyName(t, identifier);

                    }
                }

                sw.Close();

                //display
                MessageBox.Show(this, "Done", "Export done", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Error during export: " + ex.Message, "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void foreignDeviceRegistrationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //fetch end point
            BacnetClient comm = null;
            try
            {
                if (m_DeviceTree.SelectedNode == null) return;
                else if (m_DeviceTree.SelectedNode.Tag == null) return;
                else if (!(m_DeviceTree.SelectedNode.Tag is BacnetClient)) return;
                comm = (BacnetClient)m_DeviceTree.SelectedNode.Tag;
            }
            finally
            {

                if (comm == null) MessageBox.Show(this, "Please select an \"IP transport\" node first", "Wrong node", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            Form f = new ForeignRegistry(comm);
            f.ShowDialog();
        }

        private void alarmSummaryToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            alarmSummaryToolStripMenuItem_Click(sender, e);
        }

        private void alarmSummaryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //fetch end point
            BacnetClient client;
            BacnetAddress address;
            uint deviceId;

            FetchEndPoint(out client, out address, out deviceId);

            if (client == null)
            {
                MessageBox.Show(this, "Please select a device node", "Wrong node", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            new AlarmSummary(m_AddressSpaceTree.ImageList, client, address, deviceId, DevicesObjectsName).ShowDialog();
        }

        // Read the Adress Space, and change all object Id by name
        // Popup ToolTipText Get Properties Name
        private void readPropertiesNameToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //fetch end point
            BacnetClient client;
            BacnetAddress address;
            uint deviceId;

            FetchEndPoint(out client, out address, out deviceId);

            if (client == null)
            {
                MessageBox.Show(this, "Please select a device node", "Wrong node", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Go
            ChangeObjectIdByName(m_AddressSpaceTree.Nodes, client, address);

        }

        // In the Objects TreeNode, get all elements without the Bacnet PROP_OBJECT_NAME not Read out
        private void GetRequiredObjectName(TreeNodeCollection tnc, List<BacnetReadAccessSpecification> bras)
        {
            foreach (TreeNode tn in tnc)
            {
                if (tn.ToolTipText == "")
                {
                    if (!bras.Exists(o => o.objectIdentifier.Equals((BacnetObject)tn.Tag)))
                        bras.Add(new BacnetReadAccessSpecification((BacnetObject)tn.Tag, new BacnetPropertyReference[] { new BacnetPropertyReference((uint)BacnetPropertyIds.PROP_OBJECT_NAME, System.IO.BACnet.Serialize.ASN1.BACNET_ARRAY_ALL) }));
                }
                if (tn.Nodes != null)
                    GetRequiredObjectName(tn.Nodes, bras);
            }
        }
        // In the Objects TreeNode, set all elements with the ReadPropertyMultiple response
        private void SetObjectName(TreeNodeCollection tnc, IList<BacnetReadAccessResult> result, BacnetAddress adr)
        {
            foreach (TreeNode tn in tnc)
            {
                var b = (BacnetObject)tn.Tag;

                try
                {
                    if (tn.ToolTipText == "")
                    {
                        BacnetReadAccessResult r = result.Single(o => o.objectIdentifier.Equals(b));
                        ChangeTreeNodePropertyName(tn, r.values[0].value[0].ToString());
                        lock (DevicesObjectsName)
                        {
                            var t = new Tuple<string, BacnetObject>(adr.FullHashString(), (BacnetObject)tn.Tag);
                            DevicesObjectsName.Remove(t); // sometimes the same object appears at several place (in Groups for instance).
                            DevicesObjectsName.Add(t, r.values[0].value[0].ToString());
                        }
                    }
                }
                catch { }

                if (tn.Nodes != null)
                    SetObjectName(tn.Nodes, result, adr);
            }

        }
        // Try a ReadPropertyMultiple for all PROP_OBJECT_NAME not already known
        private void ChangeObjectIdByName(TreeNodeCollection tnc, BacnetClient comm, BacnetAddress adr)
        {
            var retries = comm.Retries;
            comm.Retries = 1;
            var isOk = false;

            var bras = new List<BacnetReadAccessSpecification>();
            GetRequiredObjectName(tnc, bras);

            if (bras.Count == 0)
                isOk = true;
            else
                try
                {
                    IList<BacnetReadAccessResult> result = null;
                    if (comm.ReadPropertyMultipleRequest(adr, bras, out result) == true)
                    {
                        SetObjectName(tnc, result, adr);
                        isOk = true;
                    }
                }
                catch { }

            // Fail, so go One by One, in a background thread
            if (!isOk)
                ThreadPool.QueueUserWorkItem((o) =>
                {
                    ChangeObjectIdByNameOneByOne(m_AddressSpaceTree.Nodes, comm, adr, _asyncRequestId);
                });

            comm.Retries = retries;
        }

        private void ChangeObjectIdByNameOneByOne(TreeNodeCollection tnc, BacnetClient comm, BacnetAddress adr, int asyncRequestId)
        {
            var retries = comm.Retries;
            comm.Retries = 1;

            foreach (TreeNode tn in tnc)
            {
                if (tn.ToolTipText == "")
                {
                    IList<BacnetValue> name;
                    if (comm.ReadPropertyRequest(adr, (BacnetObject)tn.Tag, BacnetPropertyIds.PROP_OBJECT_NAME, out name) == true)
                    {
                        if (asyncRequestId != _asyncRequestId) // Selected device is no more the good one
                        {
                            comm.Retries = retries;
                            return;
                        }

                        Invoke((MethodInvoker)delegate
                        {
                            if (asyncRequestId != _asyncRequestId) return; // another test in the GUI thread

                            ChangeTreeNodePropertyName(tn, name[0].Value.ToString());

                            lock (DevicesObjectsName)
                            {
                                var t = new Tuple<string, BacnetObject>(adr.FullHashString(), (BacnetObject)tn.Tag);
                                DevicesObjectsName.Remove(t); // sometimes the same object appears at several place (in Groups for instance).
                                DevicesObjectsName.Add(t, name[0].Value.ToString());
                            }
                        });
                    }
                }

                if (tn.Nodes != null)
                    ChangeObjectIdByNameOneByOne(tn.Nodes, comm, adr, asyncRequestId);

                comm.Retries = retries;
            }
        }

        // Open a serialized Dictionary of object id <-> object name file
        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //which file to upload?
            var dlg = new OpenFileDialog
            {
                FileName = Properties.Settings.Default.ObjectNameFile,
                DefaultExt = "YabeMap",
                Filter = "Yabe Map files (*.YabeMap)|*.YabeMap|All files (*.*)|*.*"
            };
            if (dlg.ShowDialog(this) != DialogResult.OK) return;
            var filename = dlg.FileName;
            Properties.Settings.Default.ObjectNameFile = filename;

            try
            {
                Stream stream = File.Open(Properties.Settings.Default.ObjectNameFile, FileMode.Open);
                var bf = new BinaryFormatter();
                var d = (Dictionary<Tuple<string, BacnetObject>, string>)bf.Deserialize(stream);
                stream.Close();

                if (d != null) DevicesObjectsName = d;
            }
            catch
            {
                MessageBox.Show(this, "File error", "Wrong file", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // save a serialized Dictionary of object id <-> object name file
        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {

            var dlg = new SaveFileDialog
            {
                FileName = Properties.Settings.Default.ObjectNameFile,
                DefaultExt = "YabeMap",
                Filter = "Yabe Map files (*.YabeMap)|*.YabeMap|All files (*.*)|*.*"
            };
            if (dlg.ShowDialog(this) != DialogResult.OK) return;
            var filename = dlg.FileName;
            Properties.Settings.Default.ObjectNameFile = filename;

            try
            {
                Stream stream = File.Open(Properties.Settings.Default.ObjectNameFile, FileMode.Create);
                var bf = new BinaryFormatter();
                bf.Serialize(stream, DevicesObjectsName);
                stream.Close();
            }
            catch
            {
                MessageBox.Show(this, "File error", "Wrong file", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        private void createObjectToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            createObjectToolStripMenuItem_Click(sender, e);
        }
        private void createObjectToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //fetch end point
            BacnetClient client;
            BacnetAddress address;
            uint deviceId;

            FetchEndPoint(out client, out address, out deviceId);

            if (client == null)
                return;

            var f = new CreateObject();
            if (f.ShowDialog() == DialogResult.OK)
            {

                try
                {
                    client.CreateObjectRequest(address, new BacnetObject((BacnetObjectTypes)f.ObjectType.SelectedIndex, (uint)f.ObjectId.Value));
                    m_DeviceTree_AfterSelect(null, new TreeViewEventArgs(m_DeviceTree.SelectedNode));
                }
                catch (Exception ex)
                {
                    Trace.TraceError("Error : " + ex.Message);
                    MessageBox.Show("Fail to Create Object", "CreateObject", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }

        }

        private void editBBMDTablesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            BacnetClient client;
            BacnetAddress address;
            uint deviceId;

            FetchEndPoint(out client, out address, out deviceId);

            if ((client != null) && (client.Transport is BacnetIpUdpProtocolTransport) && (address != null) && (address.RoutedSource == null))
                new BBMDEditor(client, address).ShowDialog();
            else
                MessageBox.Show("An IPv4 device is required", "Wrong device", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void cleanToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var res = MessageBox.Show(this, "Clean all " + DevicesObjectsName.Count + " entries, really ?", "Name database suppression", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            if (res == DialogResult.OK)
                DevicesObjectsName = new Dictionary<Tuple<string, BacnetObject>, string>();
        }

        // Change the WritePriority Value
        private void MainDialog_KeyDown(object sender, KeyEventArgs e)
        {
            if ((e.Modifiers == (Keys.Control | Keys.Alt)))
            {

                if ((e.KeyCode >= Keys.D0 && e.KeyCode <= Keys.D9) || (e.KeyCode >= Keys.NumPad0 && e.KeyCode <= Keys.NumPad9))
                {
                    var s = e.KeyCode.ToString();
                    var i = Convert.ToInt32(s[s.Length - 1]) - 48;

                    Properties.Settings.Default.DefaultWritePriority = (BacnetWritePriority)i;
                    SystemSounds.Beep.Play();
                    Trace.WriteLine("WritePriority change to level " + i + " : " + ((BacnetWritePriority)i));
                }
            }
        }

        #region "Alarm Logger"

        StreamWriter _alarmFileWriter;
        readonly object _lockAlarmFileWriter = new object();
        private bool _busy;

        private void EventAlarmLogToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (_alarmFileWriter != null)
            {
                lock (_lockAlarmFileWriter)
                {
                    _alarmFileWriter.Close();
                    EventAlarmLogToolStripMenuItem.Text = "Start saving Cov/Event/Alarm Log";
                    _alarmFileWriter = null;
                }
                return;

            }

            //which file to use ?
            var dlg = new SaveFileDialog {DefaultExt = ".csv", Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"};
            if (dlg.ShowDialog(this) != DialogResult.OK) return;
            var filename = dlg.FileName;

            try
            {
                _alarmFileWriter = new StreamWriter(filename);
                _alarmFileWriter.WriteLine("Device;ObjectId;Name;Value;Time;Status;Description");
                EventAlarmLogToolStripMenuItem.Text = "Stop saving Cov/Event/Alarm Log";
            }
            catch
            {
                MessageBox.Show(this, "File error", "Unable to open the file", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _alarmFileWriter = null;
            }
        }


        // Event/Alarm logging
        private void AddLogAlarmEvent(ListViewItem itm)
        {
            lock (_lockAlarmFileWriter)
            {
                if (_alarmFileWriter != null)
                {
                    for (var i = 0; i < itm.SubItems.Count; i++)
                    {
                        _alarmFileWriter.Write(((i != 0) ? ";" : "") + itm.SubItems[i].Text);
                    }
                    _alarmFileWriter.WriteLine();
                    _alarmFileWriter.Flush();
                }
            }
        }

        #endregion

    }

    // Used to sort the devices Tree by device_id
    public class NodeSorter : IComparer
    {
        public int Compare(object x, object y)
        {
            var tx = x as TreeNode;
            var ty = y as TreeNode;


            // Two device, compare the device_id
            if ((tx.Tag is KeyValuePair<BacnetAddress, uint> entryX) && (ty.Tag is KeyValuePair<BacnetAddress, uint> entryY))
                return entryX.Value.CompareTo(entryY.Value);
            else // something must be provide
                return tx.Text.CompareTo(ty.Text);
        }
    }
}

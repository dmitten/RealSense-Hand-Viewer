using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;

namespace hands_viewer.cs
{
    class HandsRecognition
    {
        PXCMImage.ImageInfo info;
        byte[] LUT;
        private MainForm form;
        private bool _disconnected = false;
        //Queue containing depth image - for synchronization purposes
        private Queue<PXCMImage> m_images;

        private Queue<PXCMPoint3DF32> [] m_cursorPoints;
        private int [] m_cursorClick;

        private const int NumberOfFramesToDelay = 3;
        private int _framesCounter = 0;
        private float _maxRange;

        private const int CURSOR_FACTOR_X = 60;
        private const int CURSOR_FACTOR_Y_UP = 120;
        private const int CURSOR_FACTOR_Y_DOWN = 40;

        public HandsRecognition(MainForm form)
        {
            m_images = new Queue<PXCMImage>();
            m_cursorPoints = new Queue<PXCMPoint3DF32>[2];
            m_cursorPoints[0] = new Queue<PXCMPoint3DF32>();
            m_cursorPoints[1] = new Queue<PXCMPoint3DF32>();
            m_cursorClick = new int[2];

            this.form = form;
            LUT = Enumerable.Repeat((byte)0, 256).ToArray();
            LUT[255] = 1;
        }

        //Added by Ariel (scaling Cursor rendering)

        float GetXPosition(float cursorXPos)
        {
            //cursorXPos = (cursorXPos - 1) * (-1);
            if (cursorXPos < CURSOR_FACTOR_X)
            {
                return 0.0f;
            }
            if (cursorXPos >= info.width - CURSOR_FACTOR_X)
            {
                return (float)(info.width - 1.0);
            }
            return ((cursorXPos - CURSOR_FACTOR_X) / (info.width - 2 * CURSOR_FACTOR_X)) * info.width;
        }

        float GetYPosition(float cursorYPos)
        {
            if (cursorYPos < CURSOR_FACTOR_Y_UP)
            {
                return 0.0f;
            }
            if (cursorYPos >= info.height - CURSOR_FACTOR_Y_DOWN)
            {
                return (float)(info.height - 1.0);
            }
            return ((cursorYPos - CURSOR_FACTOR_Y_UP) / (info.height - CURSOR_FACTOR_Y_UP - CURSOR_FACTOR_Y_DOWN)) * info.height;
        }


        /* Checking if sensor device connect or not */
        private bool DisplayDeviceConnection(bool state)
        {
            if (state)
            {
                if (!_disconnected) form.UpdateStatus("Device Disconnected");
                _disconnected = true;
            }
            else
            {
                if (_disconnected) form.UpdateStatus("Device Reconnected");
                _disconnected = false;
            }
            return _disconnected;
        }

        /* Displaying current frame gestures */
        private void DisplayGesture(PXCMHandData handAnalysis,int frameNumber)
        {

            int firedGesturesNumber = handAnalysis.QueryFiredGesturesNumber();
            string gestureStatusLeft = string.Empty;
            string gestureStatusRight = string.Empty;

            if (firedGesturesNumber == 0)
            {
                return;
            }
            
            for (int i = 0; i < firedGesturesNumber; i++)
            {
                PXCMHandData.GestureData gestureData;
                if (handAnalysis.QueryFiredGestureData(i, out gestureData) == pxcmStatus.PXCM_STATUS_NO_ERROR)
                {
                    PXCMHandData.IHand handData;
                    if (handAnalysis.QueryHandDataById(gestureData.handId, out handData) != pxcmStatus.PXCM_STATUS_NO_ERROR)
                        return;
                   
                    PXCMHandData.BodySideType bodySideType = handData.QueryBodySide();
                    if (bodySideType == PXCMHandData.BodySideType.BODY_SIDE_LEFT)
                    {
                        gestureStatusLeft += "Left Hand Gesture: " + gestureData.name;
                    }
                    else if (bodySideType == PXCMHandData.BodySideType.BODY_SIDE_RIGHT)
                    {
                        gestureStatusRight += "Right Hand Gesture: " + gestureData.name;
                    }
                   
                }
                  
            }
            if (gestureStatusLeft == String.Empty)
                form.UpdateInfo("Frame " + frameNumber + ") " + gestureStatusRight + "\n", Color.SeaGreen);
            else
                form.UpdateInfo("Frame " + frameNumber + ") " + gestureStatusLeft + ", " + gestureStatusRight + "\n",Color.SeaGreen);
          
        }

        /* Displaying Depth/Mask Images - for depth image only we use a delay of NumberOfFramesToDelay to sync image with tracking */
        private unsafe void DisplayPicture(PXCMImage depth, PXCMHandData handAnalysis)
        {
            if (depth == null)
                return;

            PXCMImage image = depth;

            if (form.GetCursorModeState())
            {
               
                info = image.QueryInfo();
                Bitmap depthBitmap;
                depthBitmap = new Bitmap(image.info.width, image.info.height, PixelFormat.Format8bppIndexed);
                form.DisplayBitmap(depthBitmap);
                depthBitmap.Dispose(); 
            }
            else
            {
                //Mask Image
                if (form.GetLabelmapState())
                {
                    Bitmap labeledBitmap = null;
                    try
                    {
                        int numOfHands = handAnalysis.QueryNumberOfHands();

                        PXCMPointI32[][] pointOuter = new PXCMPointI32[numOfHands][];
                        PXCMPointI32[][] pointInner = new PXCMPointI32[numOfHands][];

                        labeledBitmap = new Bitmap(image.info.width, image.info.height, PixelFormat.Format32bppRgb);
                        for (int j = 0; j < numOfHands; j++)
                        {
                            int id;
                            PXCMImage.ImageData data;

                            handAnalysis.QueryHandId(PXCMHandData.AccessOrderType.ACCESS_ORDER_BY_TIME, j, out id);
                            //Get hand by time of appearance
                            PXCMHandData.IHand handData;
                            handAnalysis.QueryHandData(PXCMHandData.AccessOrderType.ACCESS_ORDER_BY_TIME, j, out handData);
                            if (handData != null &&
                                (handData.QuerySegmentationImage(out image) >= pxcmStatus.PXCM_STATUS_NO_ERROR))
                            {
                                if (image.AcquireAccess(PXCMImage.Access.ACCESS_READ, PXCMImage.PixelFormat.PIXEL_FORMAT_Y8,
                                    out data) >= pxcmStatus.PXCM_STATUS_NO_ERROR)
                                {
                                    Rectangle rect = new Rectangle(0, 0, image.info.width, image.info.height);

                                    BitmapData bitmapdata = labeledBitmap.LockBits(rect, ImageLockMode.ReadWrite, labeledBitmap.PixelFormat);
                                    byte* numPtr = (byte*)bitmapdata.Scan0; //dst
                                    byte* numPtr2 = (byte*)data.planes[0]; //row
                                    int imagesize = image.info.width * image.info.height;
                                    byte num2 = (form.GetFullHandModeState()) ? (byte)handData.QueryBodySide() : (byte)1;

                                    byte tmp = 0;
                                    for (int i = 0; i < imagesize; i++, numPtr += 4, numPtr2++)
                                    {
                                        tmp = (byte)(LUT[numPtr2[0]] * num2 * 100);
                                        numPtr[0] = (Byte)(tmp | numPtr[0]);
                                        numPtr[1] = (Byte)(tmp | numPtr[1]);
                                        numPtr[2] = (Byte)(tmp | numPtr[2]);
                                        numPtr[3] = 0xff;
                                    }

                                    labeledBitmap.UnlockBits(bitmapdata);
                                    image.ReleaseAccess(data);

                                }

                                if ((form.GetContourState()))
                                {
                                    int contourNumber = handData.QueryNumberOfContours();
                                    if (contourNumber > 0)
                                    {
                                        for (int k = 0; k < contourNumber; ++k)
                                        {
                                            PXCMHandData.IContour contour;
                                            pxcmStatus sts = handData.QueryContour(k, out contour);
                                            if (sts == pxcmStatus.PXCM_STATUS_NO_ERROR)
                                            {
                                                //int contourSize = contour.QuerySize();
                                                if (contour.IsOuter() == true)
                                                    contour.QueryPoints(out pointOuter[j]);
                                                else
                                                {
                                                    contour.QueryPoints(out pointInner[j]);
                                                }
                                            }
                                        }

                                    }
                                }

                            }
                        }
                        if (labeledBitmap != null)
                        {

                            form.DisplayBitmap(labeledBitmap);
                            labeledBitmap.Dispose();
                        }
                        image.Dispose();

                        for (int i = 0; i < numOfHands; i++)
                        {
                            if (form.GetContourState())
                            {
                                if (pointOuter[i] != null && pointOuter[i].Length > 0)
                                    form.DisplayContour(pointOuter[i], i);
                                if (pointInner[i] != null && pointInner[i].Length > 0)
                                    form.DisplayContour(pointInner[i], i);
                            }

                        }

                    }
                    catch (Exception)
                    {
                        if (labeledBitmap != null)
                        {
                            labeledBitmap.Dispose();
                        }
                        if (image != null)
                        {
                            image.Dispose();
                        }
                    }

                }//end label image

                //Depth Image
                else
                {
                    //collecting 3 images inside a queue and displaying the oldest image
                    PXCMImage.ImageInfo info;
                    PXCMImage image2;

                    info = image.QueryInfo();
                    image2 = form.g_session.CreateImage(info);
                    if (image2 == null) { return; }
                    image2.CopyImage(image);
                    m_images.Enqueue(image2);
                    if (m_images.Count == NumberOfFramesToDelay)
                    {
                        Bitmap depthBitmap;
                        try
                        {
                            depthBitmap = new Bitmap(image.info.width, image.info.height, PixelFormat.Format32bppRgb);
                        }
                        catch (Exception)
                        {
                            image.Dispose();
                            PXCMImage queImage = m_images.Dequeue();
                            queImage.Dispose();
                            return;
                        }

                        PXCMImage.ImageData data3;
                        PXCMImage image3 = m_images.Dequeue();
                        if (image3.AcquireAccess(PXCMImage.Access.ACCESS_READ, PXCMImage.PixelFormat.PIXEL_FORMAT_DEPTH, out data3) >= pxcmStatus.PXCM_STATUS_NO_ERROR)
                        {
                            float fMaxValue = _maxRange;
                            byte cVal;

                            Rectangle rect = new Rectangle(0, 0, image.info.width, image.info.height);
                            BitmapData bitmapdata = depthBitmap.LockBits(rect, ImageLockMode.ReadWrite, depthBitmap.PixelFormat);

                            byte* pDst = (byte*)bitmapdata.Scan0;
                            short* pSrc = (short*)data3.planes[0];
                            int size = image.info.width * image.info.height;

                            for (int i = 0; i < size; i++, pSrc++, pDst += 4)
                            {
                                cVal = (byte)((*pSrc) / fMaxValue * 255);
                                if (cVal != 0)
                                    cVal = (byte)(255 - cVal);

                                pDst[0] = cVal;
                                pDst[1] = cVal;
                                pDst[2] = cVal;
                                pDst[3] = 255;
                            }
                            try
                            {
                                depthBitmap.UnlockBits(bitmapdata);
                            }
                            catch (Exception)
                            {
                                image3.ReleaseAccess(data3);
                                depthBitmap.Dispose();
                                image3.Dispose();
                                return;
                            }

                            form.DisplayBitmap(depthBitmap);
                            image3.ReleaseAccess(data3);
                        }
                        depthBitmap.Dispose();
                        image3.Dispose();
                    }
                }
            }
        }


        /* Displaying current frames hand joints */
        private void DisplayJoints(PXCMHandData handOutput, long timeStamp = 0)
        {
            m_cursorClick[0] = Math.Max(0, m_cursorClick[0] - 1);
            m_cursorClick[1] = Math.Max(0, m_cursorClick[1] - 1);
           
            if (form.GetJointsState() || form.GetSkeletonState() || form.GetCursorState() || form.GetExtremitiesState())
            {
                //Iterate hands
                PXCMHandData.JointData[][] nodes = new PXCMHandData.JointData[][] { new PXCMHandData.JointData[0x20], new PXCMHandData.JointData[0x20] };
                PXCMHandData.ExtremityData[][] extremityNodes = new PXCMHandData.ExtremityData[][] { new PXCMHandData.ExtremityData[0x6], new PXCMHandData.ExtremityData[0x6] };
               
                int numOfHands = handOutput.QueryNumberOfHands();

                if (numOfHands == 1) m_cursorPoints[1].Clear();

                for (int i = 0; i < numOfHands; i++)
                {
                    //Get hand by time of appearence
                    PXCMHandData.IHand handData;
                    if (handOutput.QueryHandData(PXCMHandData.AccessOrderType.ACCESS_ORDER_BY_TIME, i, out handData) == pxcmStatus.PXCM_STATUS_NO_ERROR)
                    {
                        if (handData != null)
                        {
                            //Iterate Joints
                            for (int j = 0; j < 0x20; j++)
                            {
                                PXCMHandData.JointData jointData;
                                handData.QueryTrackedJoint((PXCMHandData.JointType)j, out jointData);
                                nodes[i][j] = jointData;

                            } // end iterating over joints

                            // cursor
                            if (form.GetCursorModeState() && form.GetCursorState())
                            {
                                if (handData.HasCursor() == true)
                                {
                                    PXCMHandData.ICursor cursor;
                                    if (handData.QueryCursor(out cursor) >= pxcmStatus.PXCM_STATUS_NO_ERROR && cursor != null)
                                    {
                                        PXCMPoint3DF32 imagePoint = cursor.QueryPointImage();
                                        imagePoint.x = GetXPosition(imagePoint.x);
                                        imagePoint.y = GetYPosition(imagePoint.y);
                                        
                                        m_cursorPoints[i].Enqueue(imagePoint);
                                        if (m_cursorPoints[i].Count > 50)
                                            m_cursorPoints[i].Dequeue();
                                      
                                    }
                                }
                                PXCMHandData.GestureData gestureData;
                                if (handOutput.IsGestureFiredByHand("cursor_click", handData.QueryUniqueId(),out gestureData))
			                    {
				                    m_cursorClick[i] = 7;
			                    }
                            }
                            
                            // iterate over extremitiy points
                            if (form.GetExtremitiesModeState() && form.GetExtremitiesState())
                            {
				                for(int j = 0; j < PXCMHandData.NUMBER_OF_EXTREMITIES ; j++)
				                {
					                handData.QueryExtremityPoint((PXCMHandData.ExtremityType)j, out extremityNodes[i][j]);					
				                }
			                }

                        }
                    }
                } // end itrating over hands

         
                form.DisplayJoints(nodes, numOfHands);
                if (numOfHands > 0)
                {

                    if (form.GetCursorModeState() && form.GetCursorState())
                        form.DisplayCursor(numOfHands,m_cursorPoints, m_cursorClick);

                    if (form.GetExtremitiesModeState() && form.GetExtremitiesState())
                        form.DisplayExtremities(numOfHands, extremityNodes);
                }
                else
                {
                    m_cursorPoints[0].Clear();
                    m_cursorPoints[1].Clear();
                }
            }
        }

        /* Displaying current frame alerts */
        private void DisplayAlerts(PXCMHandData handAnalysis, int frameNumber)
        {
            bool isChanged = false;
            string sAlert = "Alert: ";
            for (int i = 0; i < handAnalysis.QueryFiredAlertsNumber(); i++)
            {
                PXCMHandData.AlertData alertData;
                if (handAnalysis.QueryFiredAlertData(i, out alertData) != pxcmStatus.PXCM_STATUS_NO_ERROR)
                    continue;

                //See PXCMHandAnalysis.AlertData.AlertType for all available alerts
                switch (alertData.label)
                {
                    case PXCMHandData.AlertType.ALERT_HAND_DETECTED:
                        {

                            sAlert += "Hand Detected, ";
                            isChanged = true;
                            break;
                        }
                    case PXCMHandData.AlertType.ALERT_HAND_NOT_DETECTED:
                        {

                            sAlert += "Hand Not Detected, ";
                            isChanged = true;
                            break;
                        }
                    case PXCMHandData.AlertType.ALERT_HAND_CALIBRATED:
                        {

                            sAlert += "Hand Calibrated, ";
                            isChanged = true;
                            break;
                        }
                    case PXCMHandData.AlertType.ALERT_HAND_NOT_CALIBRATED:
                        {

                            sAlert += "Hand Not Calibrated, ";
                            isChanged = true;
                            break;
                        }
                    case PXCMHandData.AlertType.ALERT_HAND_INSIDE_BORDERS:
                        {

                            sAlert += "Hand Inside Border, ";
                            isChanged = true;
                            break;
                        }
                    case PXCMHandData.AlertType.ALERT_HAND_OUT_OF_BORDERS:
                        {

                            sAlert += "Hand Out Of Borders, ";
                            isChanged = true;
                            break;
                        }
                    case PXCMHandData.AlertType.ALERT_CURSOR_DETECTED:
                        {

                            sAlert += "Cursor Detected, ";
                            isChanged = true;
                            break;
                        }
                    case PXCMHandData.AlertType.ALERT_CURSOR_NOT_DETECTED:
                        {

                            sAlert += "Cursor Not Detected, ";
                            isChanged = true;
                            break;
                        }
                    case PXCMHandData.AlertType.ALERT_CURSOR_INSIDE_BORDERS:
                        {

                            sAlert += "Cursor Inside Borders, ";
                            isChanged = true;
                            break;
                        }
                    case PXCMHandData.AlertType.ALERT_CURSOR_OUT_OF_BORDERS:
                        {

                            sAlert += "Cursor Out Of Borders, ";
                            isChanged = true;
                            break;
                        }
                }
            }
            if (isChanged)
            {
                form.UpdateInfo("Frame " + frameNumber + ") " + sAlert + "\n", Color.RoyalBlue);
            }
        }

        public static pxcmStatus OnNewFrame(Int32 mid, PXCMBase module, PXCMCapture.Sample sample)
        {
            return pxcmStatus.PXCM_STATUS_NO_ERROR;
        }


        /* Using PXCMSenseManager to handle data */
        public void SimplePipeline()
        {   
            form.UpdateInfo(String.Empty,Color.Black);
            bool liveCamera = false;
       
            bool flag = true;
            PXCMSenseManager instance = null;
            _disconnected = false;
            instance = form.g_session.CreateSenseManager();
            if (instance == null)
            {
                form.UpdateStatus("Failed creating SenseManager");
                form.EnableTrackingMode(true);
                return;
            }

            PXCMCaptureManager captureManager = instance.captureManager;
            if (captureManager != null)
            {
                if (form.GetRecordState())
                {
                    captureManager.SetFileName(form.GetFileName(), true);
                    PXCMCapture.DeviceInfo info;
                    if (form.Devices.TryGetValue(form.GetCheckedDevice(), out info))
                    {
                        captureManager.FilterByDeviceInfo(info);
                    }

                }
                else if (form.GetPlaybackState())
                {
                    captureManager.SetFileName(form.GetFileName(), false);
                   
                }
                else
                {
                    PXCMCapture.DeviceInfo info;
                    if (String.IsNullOrEmpty(form.GetCheckedDevice()))
                    {
                        form.UpdateStatus("Device Failure");
                        return;
                    }

                    if (form.Devices.TryGetValue(form.GetCheckedDevice(), out info))
                    {
                        captureManager.FilterByDeviceInfo(info);
                    }

                    liveCamera = true;
                }
            }
            /* Set Module */
            pxcmStatus status = instance.EnableHand(form.GetCheckedModule());
            PXCMHandModule handAnalysis = instance.QueryHand();

            if (status != pxcmStatus.PXCM_STATUS_NO_ERROR || handAnalysis == null)
            {
                form.UpdateStatus("Failed Loading Module");
                form.EnableTrackingMode(true);

                return;
            }

            PXCMSenseManager.Handler handler = new PXCMSenseManager.Handler();
            handler.onModuleProcessedFrame = new PXCMSenseManager.Handler.OnModuleProcessedFrameDelegate(OnNewFrame);


            PXCMHandConfiguration handConfiguration = handAnalysis.CreateActiveConfiguration();
            PXCMHandData handData = handAnalysis.CreateOutput();

            if (handConfiguration == null)
            {
                form.UpdateStatus("Failed Create Configuration");
                form.EnableTrackingMode(true);
                if (handData != null) handData.Dispose();
                if (handConfiguration != null) handConfiguration.Dispose();
                instance.Close();
                instance.Dispose();
                return;
            }
            if (handData==null)
            {
                form.UpdateStatus("Failed Create Output");
                form.EnableTrackingMode(true);
                if (handData != null) handData.Dispose();
                if (handConfiguration != null) handConfiguration.Dispose();
                instance.Close();
                instance.Dispose();
                return;
            }

            FPSTimer timer = new FPSTimer(form);
            form.UpdateStatus("Init Started");
            if (handAnalysis != null && instance.Init(handler) == pxcmStatus.PXCM_STATUS_NO_ERROR)
            {

                PXCMCapture.DeviceInfo dinfo;
                PXCMCapture.DeviceModel dModel = PXCMCapture.DeviceModel.DEVICE_MODEL_F200;
                PXCMCapture.Device device = instance.captureManager.device;
                if (device != null)
                {
                    device.QueryDeviceInfo(out dinfo);
                    dModel = dinfo.model;
                    _maxRange = device.QueryDepthSensorRange().max;

                }

                ////// cursor - F200 changes
                if (form.GetCursorModeState() && dModel == PXCMCapture.DeviceModel.DEVICE_MODEL_F200)
                {
                    form.UpdateStatus("Cursor mode is unsupported for F200 camera");
                    form.EnableTrackingMode(true);
                    if (handData != null) handData.Dispose();
                    if (handConfiguration != null) handConfiguration.Dispose();
                    instance.Close();
                    instance.Dispose();
                    return;
                }

                if (handConfiguration != null)
                {
                    PXCMHandData.TrackingModeType trackingMode = PXCMHandData.TrackingModeType.TRACKING_MODE_FULL_HAND;
			        if (form.GetCursorModeState())
                        trackingMode = PXCMHandData.TrackingModeType.TRACKING_MODE_CURSOR;
			        if (form.GetFullHandModeState())
                        trackingMode =  PXCMHandData.TrackingModeType.TRACKING_MODE_FULL_HAND;
			        if (form.GetExtremitiesModeState())
                        trackingMode =  PXCMHandData.TrackingModeType.TRACKING_MODE_EXTREMITIES;


                    handConfiguration.SetTrackingMode(trackingMode);

                    handConfiguration.EnableAllAlerts();
                    handConfiguration.EnableSegmentationImage(true);
                    bool isEnabled = handConfiguration.IsSegmentationImageEnabled();

                    handConfiguration.ApplyChanges();
                    handConfiguration.Update();

                    form.resetGesturesList();
                    int totalNumOfGestures = handConfiguration.QueryGesturesTotalNumber();
                    
                    if (totalNumOfGestures > 0)
                    {
                        this.form.UpdateGesturesToList("", 0);
                        for (int i = 0; i < totalNumOfGestures; i++)
                        {
                            string gestureName = string.Empty;
                            if (handConfiguration.QueryGestureNameByIndex(i, out gestureName) ==
                                pxcmStatus.PXCM_STATUS_NO_ERROR)
                            {
                                this.form.UpdateGesturesToList(gestureName, i + 1);
                            }
                        }

                        // form.setInitGesturesFirstTime(true);
                        form.UpdateGesturesListSize();
                    }

                   
                    // }
                }

                form.UpdateStatus("Streaming");
                int frameCounter = 0;
                int frameNumber = 0;

                while (!form.stop)
                {

                    string gestureName = form.GetGestureName();
                    if (string.IsNullOrEmpty(gestureName) == false)
                    {
                        if (handConfiguration.IsGestureEnabled(gestureName) == false)
                        {
                            handConfiguration.DisableAllGestures();
                            handConfiguration.EnableGesture(gestureName, true);
                            handConfiguration.ApplyChanges();
                        }
                    }
                    else
                    {
                        handConfiguration.DisableAllGestures();
                        handConfiguration.ApplyChanges();
                    }
                    

                    if (instance.AcquireFrame(true) < pxcmStatus.PXCM_STATUS_NO_ERROR)
                    {
                        break;
                    }

                    frameCounter++;

                    if (!DisplayDeviceConnection(!instance.IsConnected()))
                    {

                        if (handData != null)
                        {
                            handData.Update();
                        }

                        PXCMCapture.Sample sample = instance.QueryHandSample();
                        if (sample != null && sample.depth != null)
                        {
                            DisplayPicture(sample.depth, handData);

                            if (handData != null)
                            {
                                frameNumber = liveCamera ? frameCounter : instance.captureManager.QueryFrameIndex(); 

                                DisplayGesture(handData, frameNumber);
                                DisplayJoints(handData);
                                DisplayAlerts(handData, frameNumber);
                            }
                            form.UpdatePanel();
                        }
                        timer.Tick();
                    }
                    instance.ReleaseFrame();
                }
            }
            else
            {
                form.UpdateStatus("Init Failed");
                flag = false;
            }
            foreach (PXCMImage pxcmImage in m_images)
            {
                pxcmImage.Dispose();
            }

            // Clean Up
            if (handData != null) handData.Dispose();
            if (handConfiguration != null) handConfiguration.Dispose();
            instance.Close();
            instance.Dispose();

            if (flag)
            {
                form.UpdateStatus("Stopped");
            }
        }
    }
}



using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Imaging;
using MTSDKDN;

using IDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;

namespace OneNoteMathType
{
    #region MTSDK class
    class MTSDK
    {
        // c-tor
        public MTSDK() { }

        // vars
        protected bool m_bDidInit = false;

        // init
        public bool Init()
        {
            if (!m_bDidInit)
            {
                Int32 result = MathTypeSDK.Instance.MTAPIConnectMgn(MTApiStartValues.mtinitLAUNCH_AS_NEEDED, 30);
                if (result == MathTypeReturnValue.mtOK)
                {
                    m_bDidInit = true;
                    return true;
                }
                else
                    return false;
            }
            return true;
        }

        // de-init
        public bool DeInit()
        {
            if (m_bDidInit)
            {
                m_bDidInit = false;
                MathTypeSDK.Instance.MTAPIDisconnectMgn();
            }
            return true;
        }

    }
    #endregion

    #region EquationOutput Classes
    abstract class EquationOutput
    {
        // c-tor
        public EquationOutput(string strOutTrans)
        {
            if (!string.IsNullOrEmpty(strOutTrans))
                this.strOutTrans = strOutTrans;
            else
                this.strOutTrans = string.Empty;
        }

        protected EquationOutput() { }

        // properties
        protected short m_iType;
        public short iType
        {
            get { return m_iType; }
            protected set { m_iType = value; }
        }

        protected short m_iFormat;
        public short iFormat
        {
            get { return m_iFormat; }
            protected set { m_iFormat = value; }
        }

        private string m_strFileName;
        public string strFileName
        {
            get { return m_strFileName; }
            set { m_strFileName = value; }
        }

        private string m_strEquation;
        public string strEquation
        {
            get { return m_strEquation; }
            set { m_strEquation = value; }
        }

        // output translator
        protected string m_strOutTrans;
        public string strOutTrans
        {
            get { return m_strOutTrans; }
            set { m_strOutTrans = value; }
        }

        // save equation to its destination
        abstract public bool Put();
    }

    abstract class EquationOutputClipboard : EquationOutput
    {
        public EquationOutputClipboard(string strOutTrans)
            : base(strOutTrans)
        {
            strFileName = string.Empty;
            iType = MTXFormEqn.mtxfmCLIPBOARD;
        }

        public EquationOutputClipboard()
            : base()
        {
            strFileName = string.Empty;
            iType = MTXFormEqn.mtxfmCLIPBOARD;
        }

        public override bool Put() { return true; }
    }

    class EquationOutputClipboardText : EquationOutputClipboard
    {
        public EquationOutputClipboardText(string strOutTrans)
            : base(strOutTrans)
        {
            iFormat = MTXFormEqn.mtxfmTEXT;
        }

        public EquationOutputClipboardText()
            : base()
        {
            iFormat = MTXFormEqn.mtxfmTEXT;
        }

        public override string ToString() { return "Clipboard Text"; }
    }

    abstract class EquationOutputFile : EquationOutput
    {
        public EquationOutputFile(string strFileName, string strOutTrans)
            : base(strOutTrans)
        {
            this.strFileName = strFileName;
            iType = MTXFormEqn.mtxfmFILE;
        }

        protected EquationOutputFile(string strFileName)
            : base()
        {
            this.strFileName = strFileName;
            iType = MTXFormEqn.mtxfmFILE;
        }

        public override bool Put() { return true; }
    }

    class EquationOutputFileGIF : EquationOutputFile
    {
        public EquationOutputFileGIF(string strFileName)
            : base(strFileName)
        {
            iFormat = MTXFormEqn.mtxfmGIF;
        }

        public override string ToString() { return "GIF file"; }
    }

    class EquationOutputFileWMF : EquationOutputFile
    {
        public EquationOutputFileWMF(string strFileName)
            : base(strFileName)
        {
            iFormat = MTXFormEqn.mtxfmPICT;
        }

        public override string ToString() { return "WMF file"; }
    }

    class EquationOutputFileEPS : EquationOutputFile
    {
        public EquationOutputFileEPS(string strFileName)
            : base(strFileName)
        {
            iFormat = MTXFormEqn.mtxfmEPS_NONE;
        }

        public override string ToString() { return "EPS file"; }
    }

    class EquationOutputFileText : EquationOutputFile
    {
        public EquationOutputFileText(string strFileName, string strOutTrans)
            : base(strFileName, strOutTrans)
        {
            iType = MTXFormEqn.mtxfmLOCAL; // override base class as the convert function cannot directly write text files
            iFormat = MTXFormEqn.mtxfmTEXT;
        }

        public override bool Put()
        {
            try
            {
                FileStream stream = new FileStream(strFileName, FileMode.OpenOrCreate, FileAccess.Write);
                StreamWriter writer = new StreamWriter(stream);
                writer.WriteLine(strEquation);
                writer.Close();
                stream.Close();
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }

        public override string ToString() { return "Text file"; }
    }
    #endregion

    #region EquationInput Classes
    abstract class EquationInput
    {
        // c-tor
        public EquationInput(string strInTrans)
        {
            if (!string.IsNullOrEmpty(strInTrans))
                this.strInTrans = strInTrans;
            else
                this.strInTrans = string.Empty;
        }

        protected short m_iType;
        public short iType
        {
            get { return m_iType; }
            protected set { m_iType = value; }
        }

        protected short m_iFormat;
        public short iFormat
        {
            get { return m_iFormat; }
            protected set { m_iFormat = value; }
        }

        // the equation as a string
        protected string m_strEquation;
        public string strEquation
        {
            get { return m_strEquation; }
            set { m_strEquation = value; }
        }

        // the equation as a byte array
        protected byte[] m_bEquation;
        public byte[] bEquation
        {
            get { return m_bEquation; }
            set { m_bEquation = value; }
        }

        // MTEF byte array
        protected byte[] m_bMTEF;
        public byte[] bMTEF
        {
            get { return m_bMTEF; }
            set { m_bMTEF = value; }
        }

        // MTEF byte array length
        protected int m_iMTEF_Length;
        public int iMTEF_Length
        {
            get { return m_iMTEF_Length; }
            set { m_iMTEF_Length = value; }
        }

        // MTEF string
        protected string m_strMTEF;
        public string strMTEF
        {
            get { return m_strMTEF; }
            set { m_strMTEF = value; }
        }

        // input translator
        protected string m_strInTrans;
        public string strInTrans
        {
            get { return m_strInTrans; }
            set { m_strInTrans = value; }
        }

        // the source equation file
        protected string m_strFileName;
        public string strFileName
        {
            get { return m_strFileName; }
            set { m_strFileName = value; }
        }

        protected MTSDK sdk = new MTSDK();

        // get the equation from the source
        abstract public bool Get();

        // get binary MTEF
        abstract public bool GetMTEF();
    }

    abstract class EquationInputClipboard : EquationInput
    {
        public EquationInputClipboard(string strInTrans)
            : base(strInTrans)
        {
            iType = MTXFormEqn.mtxfmCLIPBOARD;
        }
    }

    class EquationInputClipboardText : EquationInputClipboard
    {
        public EquationInputClipboardText(string strInTrans)
            : base(strInTrans)
        {
            iFormat = MTXFormEqn.mtxfmTEXT;
        }

        public override bool Get() { return true; }
        public override bool GetMTEF() { return true; }
        public override string ToString() { return "Clipboard text"; }
    }

    class EquationInputClipboardEmbeddedObject : EquationInputClipboard
    {
        public EquationInputClipboardEmbeddedObject()
            : base(ClipboardFormats.cfEmbeddedObj)
        {
            iFormat = MTXFormEqn.mtxfmMTEF;
        }

        public override bool Get() { return true; }
        public override bool GetMTEF() { return true; }
        public override string ToString() { return "Clipboard Embedded Object"; }
    }

    abstract class EquationInputFile : EquationInput
    {
        public EquationInputFile(string strFileName, string strInTrans)
            : base(strInTrans)
        {
            this.strFileName = strFileName;
            iType = MTXFormEqn.mtxfmLOCAL;
        }
    }

    class EquationInputFileText : EquationInputFile
    {
        public EquationInputFileText(string strFileName, string strInTrans)
            : base(strFileName, strInTrans)
        {
            iFormat = MTXFormEqn.mtxfmMTEF;
        }

        public override string ToString() { return "Text file"; }

        override public bool Get()
        {
            try
            {
                strEquation = System.IO.File.ReadAllText(strFileName);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }

        override public bool GetMTEF()
        {
            bool bReturn = false;

            if (!sdk.Init())
                return bReturn;

            IDataObject dataObject = MathTypeSDK.getIDataObject();

            if (dataObject == null)
            {
                sdk.DeInit();
                return bReturn;
            }

            FORMATETC formatEtc = new FORMATETC();
            STGMEDIUM stgMedium = new STGMEDIUM();

            try
            {
                // Setup the formatting information to use for the conversion.
                formatEtc.cfFormat = (Int16)DataFormats.GetFormat(strInTrans).Id;
                formatEtc.dwAspect = DVASPECT.DVASPECT_CONTENT;
                formatEtc.lindex = -1;
                formatEtc.ptd = (IntPtr)0;
                formatEtc.tymed = TYMED.TYMED_HGLOBAL;

                // Setup the MathML content to convert
                stgMedium.unionmember = Marshal.StringToHGlobalAuto(strEquation);
                stgMedium.tymed = TYMED.TYMED_HGLOBAL;
                stgMedium.pUnkForRelease = 0;

                // Perform the conversion
                dataObject.SetData(ref formatEtc, ref stgMedium, false);

                // Set the format for the output
                formatEtc.cfFormat = (Int16)DataFormats.GetFormat("MathType EF").Id;
                formatEtc.dwAspect = DVASPECT.DVASPECT_CONTENT;
                formatEtc.lindex = -1;
                formatEtc.ptd = (IntPtr)0;
                formatEtc.tymed = TYMED.TYMED_ISTORAGE;

                // Create a blank data structure to hold the converted result.
                stgMedium = new STGMEDIUM();
                stgMedium.tymed = TYMED.TYMED_NULL;
                stgMedium.pUnkForRelease = 0;

                // Get the conversion result in MTEF format
                dataObject.GetData(ref formatEtc, out stgMedium);
            }
            catch (COMException e)
            {
                Console.WriteLine("MathML conversion to MathType threw an exception: " + Environment.NewLine + e.ToString());
                sdk.DeInit();
                return bReturn;
            }

            // The pointer now becomes a Handle reference.
            HandleRef handleRef = new HandleRef(null, stgMedium.unionmember);

            try
            {
                // Lock in the handle to get the pointer to the data
                IntPtr ptrToHandle = MathTypeSDK.GlobalLock(handleRef);

                // Get the size of the memory block
                m_iMTEF_Length = MathTypeSDK.GlobalSize(handleRef);

                // New an array of bytes and Marshal the data across.
                m_bMTEF = new byte[m_iMTEF_Length];
                Marshal.Copy(ptrToHandle, m_bMTEF, 0, m_iMTEF_Length);
                m_strMTEF = System.Text.ASCIIEncoding.ASCII.GetString(m_bMTEF);
                bReturn = true;
            }
            catch (Exception e)
            {
                Console.WriteLine("Generation of image from MathType failed: " + Environment.NewLine + e.ToString());
            }
            finally
            {
                MathTypeSDK.GlobalUnlock(handleRef);
            }

            sdk.DeInit();
            return bReturn;
        }
    }

    class EquationInputFileWMF2 : EquationInputFile
    {
        public EquationInputFileWMF2(string strFileName)
            : base(strFileName, "")
        {
            iFormat = MTXFormEqn.mtxfmEPS_WMF;
            iType = MTXFormEqn.mtxfmFILE;
        }

        public override bool Get() { return true; }

        public override bool GetMTEF() { return true; }
    }

    class EquationInputFileWMF : EquationInputFile
    {
        public EquationInputFileWMF(string strFileName)
            : base(strFileName, "")
        {
            iFormat = MTXFormEqn.mtxfmMTEF;
        }

        public override bool Get() { return true; }

        public override string ToString() { return "WMF file"; }

        public override bool GetMTEF()
        {
            Play();
            if (!Succeeded())
                return false;
            return true;
        }

        protected class WmfForm : Form
        {
            public WmfForm() { }
        }
        protected WmfForm wf = new WmfForm();

        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        protected struct wmfHeader
        {
            public Int16 iComment;
            public Int16 ix1;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 8)]
            public string strSig;
            public Int16 iVer;
            public Int32 iTotalLen;
            public Int32 iDataLen;
        };
        protected wmfHeader m_wmfHeader;

        protected Metafile m_metafile;
        protected const string m_strSig = "AppsMFC";
        protected bool m_succeeded = false;

        protected void Play()
        {
            try
            {
                m_succeeded = false;
                Graphics.EnumerateMetafileProc metafileDelegate;
                Point destPoint;
                m_metafile = new Metafile(strFileName);
                metafileDelegate = new Graphics.EnumerateMetafileProc(MetafileCallback);
                destPoint = new Point(20, 10);
                Graphics graphics = wf.CreateGraphics();
                graphics.EnumerateMetafile(m_metafile, destPoint, metafileDelegate);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        protected bool Succeeded() { return m_succeeded; }

        protected bool MetafileCallback(
            EmfPlusRecordType recordType,
            int flags,
            int dataSize,
            IntPtr data,
            PlayRecordCallback callbackData)
        {
            byte[] dataArray = null;
            if (data != IntPtr.Zero)
            {
                dataArray = new byte[dataSize];
                Marshal.Copy(data, dataArray, 0, dataSize);
                if (recordType == EmfPlusRecordType.WmfEscape && dataSize >= Marshal.SizeOf(m_wmfHeader) && !m_succeeded)
                {
                    m_wmfHeader = (wmfHeader)RawDeserialize(dataArray, 0, m_wmfHeader.GetType());
                    if (m_wmfHeader.strSig.Equals(m_strSig, StringComparison.CurrentCultureIgnoreCase))
                    {
                        System.Text.ASCIIEncoding enc = new System.Text.ASCIIEncoding();
                        string strCompanyInfo = enc.GetString(dataArray, Marshal.SizeOf(m_wmfHeader), m_wmfHeader.iDataLen);
                        int iNull = strCompanyInfo.IndexOf('\0');
                        if (iNull >= 0)
                        {
                            int mtefStart = Marshal.SizeOf(m_wmfHeader) + iNull + 1;
                            iMTEF_Length = m_wmfHeader.iDataLen;
                            bMTEF = new byte[iMTEF_Length];
                            Array.Copy(dataArray, mtefStart, bMTEF, 0, iMTEF_Length);
                            m_succeeded = true;
                        }
                    }
                }
            }

            m_metafile.PlayRecord(recordType, flags, dataSize, dataArray);

            return true;
        }

        protected static object RawDeserialize(byte[] rawData, int position, Type anyType)
        {
            int rawsize = Marshal.SizeOf(anyType);
            if (rawsize > rawData.Length)
                return null;
            IntPtr buffer = Marshal.AllocHGlobal(rawsize);
            Marshal.Copy(rawData, position, buffer, rawsize);
            object retobj = Marshal.PtrToStructure(buffer, anyType);
            Marshal.FreeHGlobal(buffer);
            return retobj;
        }
    }

    class EquationInputFileGIF : EquationInputFile
    {
        public EquationInputFileGIF(string strFileName)
            : base(strFileName, "")
        {
            iFormat = MTXFormEqn.mtxfmMTEF;
        }

        public override string ToString() { return "GIF file"; }

        override public bool Get()
        {
            try
            {
                FileStream stream = new FileStream(strFileName, FileMode.Open, FileAccess.Read);
                BinaryReader reader = new BinaryReader(stream);
                int iArrayLength = (Int32)stream.Length;
                bEquation = reader.ReadBytes(iArrayLength);
                reader.Close();
                stream.Close();
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }

        byte[] signature = { 0x21, 0xFF, 0x0B, 0x4d, 0x61, 0x74, 0x68, 0x54, 0x79, 0x70, 0x65, 0x30, 0x30, 0x31 };

        /*
        extracting MTEF from GIF files is described in the MathType SDK documentation, by default installed here:
        file:///C:/Program%20Files/MathTypeSDK/SDK/docs/MTEFstorage.htm#GIF%20Image%20Files
        */
        public override bool GetMTEF()
        {
            try
            {
                // search for signature
                int iSigStart = 0;
                while ((iSigStart = Array.IndexOf(bEquation, signature[0], iSigStart)) >= 0)
                {
                    if (CompareArrays(bEquation, iSigStart, bEquation.Length, signature, 0, signature.Length))
                    {
                        // signature found
                        int iIndex = iSigStart + signature.Length; // source array index
                        iMTEF_Length = 0;						   // destination array index
                        byte bLen;								   // current block length

                        try
                        {
                            // copy MTEF blocks
                            while ((bLen = (byte)bEquation.GetValue(iIndex)) > 0)
                            {
                                // resize destination array
                                Array.Resize(ref m_bMTEF, iMTEF_Length + bLen);
                                // copy from source to destination array
                                Array.Copy(bEquation, iIndex + 1, bMTEF, iMTEF_Length, bLen);
                                // track length of destination array
                                iMTEF_Length += bLen;
                                // point to next block in source array
                                iIndex += bLen + 1;
                            }
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message);
                            return false;
                        }
                        return true;
                    }
                    iSigStart++;
                }

                return false;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }

        protected bool CompareArrays(Array left, int leftStart, int leftLen, Array right, int rightStart, int rightLen)
        {
            int leftCompareNum = leftLen - leftStart;
            int rightCompareNum = rightLen - rightStart;
            int compareNum = leftCompareNum > rightCompareNum ? rightCompareNum : leftCompareNum;

            for (int x = 0; x < compareNum; x++)
            {
                if ((byte)left.GetValue(leftStart + x) != (byte)right.GetValue(rightStart + x))
                    return false;
            }
            return true;
        }
    }

    class EquationInputFileEPS : EquationInputFile
    {
        public EquationInputFileEPS(string strFileName)
            : base(strFileName, "")
        {
            iFormat = MTXFormEqn.mtxfmTEXT;
        }

        public override string ToString() { return "EPS file"; }

        public override bool Get()
        {
            try
            {
                strEquation = System.IO.File.ReadAllText(strFileName);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }

        public override bool GetMTEF()
        {
            /*
            extracting MTEF from EPS files is described in the MathType SDK documentation, by default installed here:
            file:///C:/Program%20Files/MathTypeSDK/SDK/docs/MTEFstorage.htm#Translator%20Output
            */
            const string strSig1 = "MathType";
            const string strSig2 = "MTEF";
            int iSig1Start = 0;
            while ((iSig1Start = strEquation.IndexOf(strSig1, iSig1Start)) >= 0)
            {
                int iSig2Start = strEquation.IndexOf(strSig2, iSig1Start + 1);
                int iDelimStart = iSig1Start + strSig1.Length;
                int iDelimLen = iSig2Start - iDelimStart;
                if (iSig2Start < 0 || iDelimLen != 1)
                {
                    iSig1Start++;
                    continue;
                }
                string strDelim = strEquation.Substring(iDelimStart, iDelimLen);
                int id1 = strEquation.IndexOf(strDelim, iSig1Start);
                int id2 = strEquation.IndexOf(strDelim, id1 + 1);
                int id3 = strEquation.IndexOf(strDelim, id2 + 1);
                int id4 = strEquation.IndexOf(strDelim, id3 + 1);
                int id5 = strEquation.IndexOf(strDelim, id4 + 1);
                int id6 = strEquation.IndexOf(strDelim, id5 + 1);
                m_strMTEF = strEquation.Substring(iSig1Start, id6 - iSig1Start + 1);
                bMTEF = System.Text.Encoding.ASCII.GetBytes(m_strMTEF);
                iMTEF_Length = bMTEF.Length;
                return true;
            }
            return false;
        }
    }
    #endregion

    #region ConvertEquation Class
    class ConvertEquation
    {
        protected EquationInput m_ei;
        protected EquationOutput m_eo;
        protected MTSDK m_sdk = new MTSDK();

        // c-tor
        public ConvertEquation() { }

        // convert
        virtual public bool Convert(EquationInput ei, EquationOutput eo)
        {
            m_ei = ei;
            m_eo = eo;
            return Convert();
        }

        virtual protected bool Convert()
        {
            bool bReturn = false;

            Console.WriteLine("Converting {0} to {1}", m_ei.ToString(), m_eo.ToString());

            Console.WriteLine("Get equation: {0}", m_ei.strFileName);
            if (m_ei.Get())
            {
                Console.WriteLine("Get MTEF");
                if (m_ei.GetMTEF())
                {
                    Console.WriteLine("Convert Equation");
                    if (ConvertToOutput())
                    {
                        Console.WriteLine("Write equation: {0}", m_eo.strFileName);
                        if (m_eo.Put())
                            bReturn = true;
                    }
                }
            }

            Console.WriteLine("Convert success: {0}\r\n", bReturn.ToString());
            return bReturn;
        }

        protected bool SetTranslator()
        {
            if (string.IsNullOrEmpty(m_eo.strOutTrans))
                return true;

            Int32 stat = MathTypeSDK.Instance.MTXFormSetTranslatorMgn(
                MTXFormSetTranslator.mtxfmTRANSL_INC_NAME + MTXFormSetTranslator.mtxfmTRANSL_INC_DATA,
                m_eo.strOutTrans);
            return stat == MathTypeReturnValue.mtOK;
        }

        protected bool ConvertToOutput()
        {
            bool bResult = false;
            try
            {
                if (!m_sdk.Init())
                    return false;

                if (MathTypeSDK.Instance.MTXFormResetMgn() == MathTypeReturnValue.mtOK &&
                    SetTranslator())
                {
                    Int32 stat = 0;
                    Int32 iBufferLength = 5000;
                    StringBuilder strDest = new StringBuilder(iBufferLength);
                    MTAPI_DIMS dims = new MTAPI_DIMS();

                    // convert
                    stat = MathTypeSDK.Instance.MTXFormEqnMgn(
                        m_ei.iType,
                        m_ei.iFormat,
                        m_ei.bMTEF,
                        m_ei.iMTEF_Length,
                        m_eo.iType,
                        m_eo.iFormat,
                        strDest,
                        iBufferLength,
                        m_eo.strFileName,
                        ref dims);

                    // save equation
                    if (stat == MathTypeReturnValue.mtOK)
                    {
                        m_eo.strEquation = strDest.ToString();
                        bResult = true;
                    }
                }

                m_sdk.DeInit();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return bResult;
        }
    }
    #endregion  
}

#region MTXFormEqn Doc
/*
	SHORT		src,			// [in] Equation data source, either 
								//  mtxfmPREVIOUS => data from previous result
								//  mtxfmCLIPBOARD => data on clipboard
								//  mtxfmLOCAL => data passed (i.e. in srcData)
	SHORT		srcFmt,			// [in] Equation source data format (mtxfmXXX, see next)
								//	 Note: srcFmt, srcData, and srcDataLen are used only
								//		if src is mtfxmLOCAL
	LPCVOID		srcData,		// [in] Depends on data source (src)
								//  mtxfmMTEF => ptr to MTEF-binary (BYTE *)
								//  mtxfmPICT => ptr to pict (MTAPI_PICT *)
								//  mtxfmTEXT => ptr to text (CHAR *), either MTEF-text or plain text
	LONG		srcDataLen,		// [in] # of bytes in srcData
 * 
 * ///////////////////////////////
 * 
	SHORT		dst,		    // [in] Equation data destination, either
								//  mtxfmCLIPBOARD => transformed data placed on clipboard
								//  mtxfmLOCAL => transformed data in dstData
								//  mtxfmFILE => transformed data in the file specified by dstPath
	SHORT		dstFmt,			// [in] Equation data format (mtxfmXXX, see next)
								//	 Note: dstFmt, dstData, and dstDataLen are used only
								//		if dst is mtfxmLOCAL (data placed on the clipboard 
								//		is either an OLE object or translator text)
	LPVOID		dstData,		// [out] Depends on data destination (dstFmt)
								//  mtxfmMTEF => ptr to MTEF-binary (BYTE *)
								//  mtxfmHMTEF => ptr to handle to MTEF-binary (HANDLE *)
								//  mtxfmPICT => ptr to pict data (MTAPI_PICT *)
								//  mtxfmTEXT => ptr to translated text or, if no translator, MTEF-text (CHAR *)
								//  mtxfmHTEXT => ptr to handle to translated text or, if no translator, MTEF-text (HANDLE *)
								//  Note: If translator specified dst must be either
								//		mtxfmTEXT or mtxfmHTEXT for the translation to be performed
	LONG		dstDataLen,		// [in] # of bytes in dstData (used for mtxfmLOCAL only)
	LPCSTR		dstPath,		// [in] destination pathname (used if dst == mtxfmFILE only, may be NULL if not used)
 * 
 * ///////////////////////////////
 * 
	MTAPI_DIMS *dims			// [out] pict dimensions, may be NULL (valid only for 
								//		dst = mtxfmPICT)
*/
#endregion

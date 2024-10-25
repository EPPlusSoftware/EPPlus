namespace OfficeOpenXml.Drawing.OleObject
{
    /// <summary>
    /// Object containing additional parameters for OLE Objects.
    /// </summary>
    public class ExcelOleObjectParameters
    {
        /// <summary>
        /// file path for ole object.
        /// </summary>
        internal string OlePath = null;
        /// <summary>
        /// True: File will be linked. False: File will be embedded.
        /// </summary>
        public bool LinkToFile = false;
        /// <summary>
        /// Set to display the object as in icon.
        /// </summary>
        public bool DisplayAsIcon = false;
        /// <summary>
        /// Use to set custom progId.
        /// </summary>
        public string ProgId = null;
        private string _extension;
        /// <summary>
        /// File Extension of OLE Object.
        /// </summary>
        public string Extension
        {
            get
            {
                return _extension;
            }
            set
            {
                _extension = value.StartsWith(".") ? value : "." + value;
            }
        }
    }
}

namespace OfficeOpenXml.Drawing.EMF
{
    /// <summary>
    /// Defines maping of Page Space or Logical Units (LU) into device space. 
    /// </summary>
    internal enum MapMode
    {
        /// <summary>
        /// LU mapped to one device pixel. positive X = Right. positive Y = down
        /// </summary>
        MM_TEXT = 0x01,
        /// <summary>
        /// LU = 0.1mm  positive X = Right. positive Y = up
        /// </summary>
        MM_LOMETRIC = 0x02,
        /// <summary>
        /// LU = 0.01mm  positive X = Right. positive Y = up
        /// </summary>
        MM_HIMETRIC = 0x03,
        /// <summary>
        /// LU = 0.01 inch positive X = Right. positive Y = up
        /// </summary>
        MM_LOENGLISH = 0x04,
        /// <summary>
        /// LU = 0.001 inch positive X = Right. positive Y = up
        /// </summary>
        MM_HIENGLISH = 0x05,
        /// <summary>
        /// LU = 1/20 printer point AKA 1/1440 inch or a ('twip'). Positive X = Right, positive Y = up
        /// </summary>
        MM_TWIPS = 0x06,
        /// <summary>
        /// EMR_SETWINDOWEXTEX and EMR_SETVIEWPORTEXTEX specify the units and orientation of axises.
        /// Equally scaled(isotropic) abstract axises
        /// X and Y must be adjusted so that they remain the same size. 
        /// When window extent is set the viewport must be adjusted to stay isotropic
        /// </summary>
        MM_ISOTROPIC = 0x07,
        /// <summary>
        /// EMR_SETWINDOWEXTEX and EMR_SETVIEWPORTEXTEX specify the units, orientation and scaling of axises.
        /// </summary>
        MM_ANISOTROPIC = 0x08
    }
}

/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/

using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Utils;
using System;
using System.IO;
#if !NET35 && !NET40
using System.Threading.Tasks;
#endif

#if NETFULL
using System.Drawing;
#endif
namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Represents an image 
    /// </summary>
    public class ExcelImage
    {
        IPictureContainer _container;
        ePictureType[] _restrictedTypes = new ePictureType[0];

        internal ExcelImage(IPictureContainer container, ePictureType[] restrictedTypes = null)
        {
            this._container = container;

            if (restrictedTypes != null)
            {
                this._restrictedTypes = restrictedTypes;
            }
        }

        /// <summary>
        /// Creates an ExcelImage to be used as template for adding images.
        /// </summary>
        public ExcelImage()
        {
        }

        /// <summary>
        /// Creates an ExcelImage to be used as template for adding images.
        /// </summary>
        /// <param name="imagePath">A path to the image file to load</param>
        public ExcelImage(string imagePath)
        {
            this.SetImage(imagePath);
        }

        /// <summary>
        /// Creates an ExcelImage to be used as template for adding images.
        /// </summary>
        /// <param name="imageFile">A FileInfo referencing the image file to load</param>
        public ExcelImage(FileInfo imageFile)
        {
            this.SetImage(imageFile);
        }

        /// <summary>
        /// Creates an ExcelImage to be used as template for adding images.
        /// </summary>
        /// <param name="imageStream">The stream containing the image</param>
        /// <param name="pictureType">The type of image loaded in the stream</param>
        public ExcelImage(Stream imageStream, ePictureType pictureType)
        {
            this.SetImage(imageStream, pictureType);
        }

        /// <summary>
        /// Creates an ExcelImage to be used as template for adding images.
        /// </summary>
        /// <param name="imageBytes">The image as a byte array</param>
        /// <param name="pictureType">The type of image loaded in the stream</param>
        public ExcelImage(byte[] imageBytes, ePictureType pictureType)
        {
            this.SetImage(imageBytes, pictureType);
        }

        /// <summary>
        /// If this object contains an image.
        /// </summary>
        public bool HasImage
        {
            get { return this.Type.HasValue; }
        }

        /// <summary>
        /// The type of image.
        /// </summary>
        public ePictureType? Type { get; internal set; }

        /// <summary>
        /// The image as a byte array.
        /// </summary>
        public byte[] ImageBytes { get; internal set; }

        /// <summary>
        /// The image bounds and resolution
        /// </summary>
        public ExcelImageInfo Bounds { get; internal set; } = new ExcelImageInfo();

        /// <summary>
        /// Sets a new image. 
        /// </summary>
        /// <param name="imagePath">The path to the image file.</param>
        public void SetImage(string imagePath)
        {
            if (string.IsNullOrEmpty(imagePath))
            {
                throw new ArgumentNullException(nameof(imagePath), "Image Path cannot be empty");
            }

            FileInfo? fi = new FileInfo(imagePath);

            if (fi.Exists == false)
            {
                throw new FileNotFoundException(imagePath);
            }

            ePictureType type = PictureStore.GetPictureType(fi.Extension);
            this.SetImage(File.ReadAllBytes(imagePath), type, true);
        }

        /// <summary>
        /// Sets a new image. 
        /// </summary>
        /// <param name="imageFile">The image file.</param>
        public void SetImage(FileInfo imageFile)
        {
            if (imageFile == null)
            {
                throw new ArgumentNullException(nameof(imageFile), "ImageFile cannot be null");
            }

            if (imageFile.Exists == false)
            {
                throw new FileNotFoundException(imageFile.FullName);
            }

            ePictureType type = PictureStore.GetPictureType(imageFile.Extension);
            this.SetImage(File.ReadAllBytes(imageFile.FullName), type, true);
        }

        /// <summary>
        /// Sets a new image. 
        /// </summary>
        /// <param name="imageBytes">The image as a byte array.</param>
        /// <param name="pictureType">The type of image.</param>
        public ExcelImage SetImage(byte[] imageBytes, ePictureType pictureType)
        {
            return this.SetImage(imageBytes, pictureType, true);
        }

        /// <summary>
        /// Sets a new image. 
        /// </summary>
        /// <param name="image">The image object to use.</param>
        /// <seealso cref="ExcelImage"/>
        public ExcelImage SetImage(ExcelImage image)
        {
            if (image.Type == null)
            {
                throw new ArgumentNullException("Image type must not be null");
            }

            return this.SetImage(image.ImageBytes, image.Type.Value, true);
        }

        /// <summary>
        /// Sets a new image. 
        /// </summary>
        /// <param name="imageStream">The stream containing the image.</param>
        /// <param name="pictureType">The type of image.</param>
        public ExcelImage SetImage(Stream imageStream, ePictureType pictureType)
        {
            if (imageStream is MemoryStream ms)
            {
                return this.SetImage(ms.ToArray(), pictureType, true);
            }
            else
            {
                if (imageStream.CanRead == false || imageStream.CanSeek == false)
                {
                    throw new ArgumentException("Stream must be readable and seekble", nameof(imageStream));
                }

                byte[]? byRet = new byte[imageStream.Length];
                imageStream.Seek(0, SeekOrigin.Begin);
                imageStream.Read(byRet, 0, (int)imageStream.Length);

                return this.SetImage(byRet, pictureType);
            }
        }
#if !NET35
        /// <summary>
        /// Sets a new image. 
        /// </summary>
        /// <param name="imageStream">The stream containing the image.</param>
        /// <param name="pictureType">The type of image.</param>
        public async Task<ExcelImage> SetImageAsync(Stream imageStream, ePictureType pictureType)
        {
            if (imageStream is MemoryStream ms)
            {
                return this.SetImage(ms.ToArray(), pictureType, true);
            }
            else
            {
                if (imageStream.CanRead == false || imageStream.CanSeek == false)
                {
                    throw new ArgumentException("Stream must be readable and seekble", nameof(imageStream));
                }

                byte[]? byRet = new byte[imageStream.Length];
                imageStream.Seek(0, SeekOrigin.Begin);
                await imageStream.ReadAsync(byRet, 0, (int)imageStream.Length);

                return this.SetImage(byRet, pictureType);
            }
        }

        /// <summary>
        /// Sets a new image. 
        /// </summary>
        /// <param name="imagePath">The path to the image file.</param>
        public async Task<ExcelImage> SetImageAsync(string imagePath)
        {
            if (string.IsNullOrEmpty(imagePath))
            {
                throw new ArgumentNullException(nameof(imagePath), "Image Path cannot be empty");
            }

            FileInfo? fi = new FileInfo(imagePath);

            return await this.SetImageAsync(fi);
        }

        /// <summary>
        /// Sets a new image. 
        /// </summary>
        /// <param name="imageFile">The image file.</param>
        public async Task<ExcelImage> SetImageAsync(FileInfo imageFile)
        {
            if (imageFile == null)
            {
                throw new ArgumentNullException(nameof(imageFile), "ImageFile cannot be null");
            }

            if (imageFile.Exists == false)
            {
                throw new FileNotFoundException(imageFile.FullName);
            }

            ePictureType type = PictureStore.GetPictureType(imageFile.Extension);
            FileStream? fs = imageFile.OpenRead();
            byte[]? b = new byte[fs.Length];
            await fs.ReadAsync(b, 0, b.Length);

            return this.SetImage(b, type, true);
        }

#endif
        internal ExcelImage SetImageNoContainer(byte[] image, ePictureType pictureType)
        {
            this.Type = pictureType;

            if (pictureType == ePictureType.Wmz || pictureType == ePictureType.Emz)
            {
                byte[]? img = ImageReader.ExtractImage(image, out ePictureType? pt);

                if (pt.HasValue)
                {
                    throw new ArgumentException($"Image is not of type {pictureType}.", nameof(image));
                }
                else
                {
                    this.ImageBytes = img;
                    pictureType = pt.Value;
                }
            }
            else
            {
                this.ImageBytes = image;
            }

            MemoryStream? ms = RecyclableMemory.GetStream(image);
            GenericImageHandler? imageHandler = new GenericImageHandler();

            if (imageHandler.GetImageBounds(ms,
                                            pictureType,
                                            out double height,
                                            out double width,
                                            out double horizontalResolution,
                                            out double verticalResolution))
            {
                this.Bounds.Width = width;
                this.Bounds.Height = height;
                this.Bounds.HorizontalResolution = horizontalResolution;
                this.Bounds.VerticalResolution = verticalResolution;
            }
            else
            {
                throw new InvalidOperationException($"The image format is not supported: {pictureType} or the image is corrupt ");
            }

            return this;
        }

        internal ExcelImage SetImage(byte[] image, ePictureType pictureType, bool removePrevImage)
        {
            if (this._container == null)
            {
                return this.SetImageNoContainer(image, pictureType);
            }
            else
            {
                return this.SetImageContainer(image, pictureType, removePrevImage);
            }
        }

        private ExcelImage SetImageContainer(byte[] image, ePictureType pictureType, bool removePrevImage)
        {
            this.ValidatePictureType(pictureType);
            this.Type = pictureType;

            if (pictureType == ePictureType.Wmz || pictureType == ePictureType.Emz)
            {
                byte[]? img = ImageReader.ExtractImage(image, out ePictureType? pt);

                if (pt.HasValue)
                {
                    throw new ArgumentException($"Image is not of type {pictureType}.", nameof(image));
                }
                else
                {
                    if (string.IsNullOrEmpty(this._container.ImageHash) == false && removePrevImage)
                    {
                        this.RemoveImageContainer();
                    }

                    this.ImageBytes = img;
                    pictureType = pt.Value;
                }
            }
            else
            {
                if (removePrevImage && string.IsNullOrEmpty(this._container.ImageHash) == false)
                {
                    this.RemoveImageContainer();
                }

                this.ImageBytes = image;
            }

            PictureStore.SavePicture(image, this._container, pictureType);
            MemoryStream? ms = RecyclableMemory.GetStream(image);

            if (this._container.RelationDocument.Package.Settings.ImageSettings.GetImageBounds(ms,
                                                                                               pictureType,
                                                                                               out double height,
                                                                                               out double width,
                                                                                               out double horizontalResolution,
                                                                                               out double verticalResolution))
            {
                this.Bounds.Width = width;
                this.Bounds.Height = height;
                this.Bounds.HorizontalResolution = horizontalResolution;
                this.Bounds.VerticalResolution = verticalResolution;
            }
            else
            {
                throw new InvalidOperationException($"Image format not supported or: {pictureType} or corrupt image");
            }

            this._container.SetNewImage();

            return this;
        }

        private void ValidatePictureType(ePictureType pictureType)
        {
            if (Array.Exists(this._restrictedTypes, x => x == pictureType))
            {
                throw new InvalidOperationException($"Picture type {pictureType} is not supported for this operation.");
            }
        }

        internal void RemoveImage()
        {
            this.RemoveImageContainer();
            this.ImageBytes = null;
            this.Type = null;
            this.Bounds = new ExcelImageInfo();
        }

        private void RemoveImageContainer()
        {
            this._container.RemoveImage();
            this._container.RelPic = null;
            this._container.ImageHash = null;
            this._container.UriPic = null;
        }
    }
}
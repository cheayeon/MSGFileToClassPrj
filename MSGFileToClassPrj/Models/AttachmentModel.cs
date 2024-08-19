﻿using MSGFileToClassPrj.Enviroment;
using System;

namespace MSGFileToClassPrj.Models
{
    public class AttachmentModel : MSGFileReadBaseModel
    {
        #region Property(s)

        private string _filename;
        /// <summary>
        /// Gets the filename.
        /// </summary>
        /// <value>The filename.</value>
        public string Filename
        {
            get
            {
                return _filename;
            }
        }

        private byte[] _data;
        /// <summary>
        /// Gets the data.
        /// </summary>
        /// <value>The data.</value>
        public byte[] Data
        {
            get { return _data; }
        }

        private object _dataObs;
        public object DataObs
        {
            get { return _dataObs; }
        }

        private string _contentId;
        /// <summary>
        /// Gets the content id.
        /// </summary>
        /// <value>The content id.</value>
        public string ContentId
        {
            get { return _contentId; }
        }

        private int _renderingPosisiton;
        /// <summary>
        /// Gets the rendering posisiton.
        /// </summary>
        /// <value>The rendering posisiton.</value>
        public int RenderingPosisiton
        {
            get { return _renderingPosisiton; }
        }

        #endregion

        public AttachmentModel(MSGFileReadBaseModel msgFileReadBaseModel) 
                : base(msgFileReadBaseModel.storage)
        {
            this.propHeaderSize = MSGFileEnv.PROPERTIES_STREAM_HEADER_ATTACH_OR_RECIP;
            this.SetMapiThisProperty();
        }

        public override void SetMapiThisProperty()
        {
            AttachmentModel setModel = this;

            foreach (var stream in streamStatistics)
            {
                if (stream.Key.StartsWith(MSGFileEnv.DATA_STORAGE))
                {
                    //FileMessageTag propTag = (FileMessageTag)Int32.Parse(stream.Key.Substring(12, 8), System.Globalization.NumberStyles.HexNumber);
                    string tagnum = stream.Key.Substring(12, 4);
                    MSGFileEnv.FileAttachmentTag propTag = (MSGFileEnv.FileAttachmentTag)Int32.Parse(tagnum, System.Globalization.NumberStyles.HexNumber);
                    NativeCOMMethods.OutLookMAPI propType = (NativeCOMMethods.OutLookMAPI)ushort.Parse(stream.Key.Substring(16, 4), System.Globalization.NumberStyles.HexNumber);

                    byte[] getData = GetStreamBytes(stream.Value);

                    switch (propTag)
                    {
                        case MSGFileEnv.FileAttachmentTag.PR_ATTACHMENT_LINKID:
                            this._contentId = ByteToString(getData, propType);
                            break;
                        case MSGFileEnv.FileAttachmentTag.PR_ATTACH_CONTENT_ID:
                            this._contentId = ByteToString(getData, propType);
                            break;
                        case MSGFileEnv.FileAttachmentTag.PR_ATTACH_DATA:
                            this._data = getData;
                            break;
                        case MSGFileEnv.FileAttachmentTag.PR_ATTACH_FILENAME:
                            this._filename = ByteToString(getData, propType);
                            break;
                        case MSGFileEnv.FileAttachmentTag.PR_ATTACH_LONG_FILENAME:
                            this._filename = ByteToString(getData, propType);
                            break;
                        case MSGFileEnv.FileAttachmentTag.PR_ATTACH_METHOD:
                            this._filename = ByteToString(getData, propType);
                            break;
                        case MSGFileEnv.FileAttachmentTag.PR_RENDERING_POSITION:
                            this._renderingPosisiton = int.Parse(ByteToString(getData, propType));
                            break;
                        default:
                            datasDictionary.Add(tagnum, getData);
                            encodingTypeDictionary.Add(tagnum, propType);
                            break;
                    }
                }
                else if (stream.Key.StartsWith(MSGFileEnv.PROPERTIES_STREAM))// property에 데이터가 있는경우는 따로 또 찾아줘야함.
                {
                    byte[] getData = this.GetStreamBytes(stream.Value);
                    datasDictionary.Add(stream.Key, getData);
                }

            }
        }
    }
}
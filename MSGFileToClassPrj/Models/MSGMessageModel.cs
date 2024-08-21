using MSGFileToClassPrj.Enviroment;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ComTypes = System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace MSGFileToClassPrj.Models
{
    public class MSGMessageModel : MSGFileReadBaseModel
    {
        #region 최상위 확인 property
        /// <summary>
        /// A reference to the parent message that this message may belong to.
        /// </summary>
        public MSGMessageModel parentMessage = null;

        /// <summary>
        /// Gets a value indicating whether this instance is the top level outlook message.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if this instance is the top level outlook message; otherwise, <c>false</c>.
        /// </value>
        public bool IsTopParent
        {
            get
            {
                if (this.parentMessage != null)
                {
                    return false;
                }
                return true;
            }
        }

        /// <summary>
        /// Gets the top level outlook message from a sub message at any level.
        /// </summary>
        /// <value>The top level outlook message.</value>
        public MSGFileReadBaseModel TopParent
        {
            get
            {
                if (this.parentMessage != null)
                {
                    return this.parentMessage.TopParent;
                }
                return this;
            }
        }
        #endregion

        #region Property(s)
        /// <summary>
        /// 아웃룩 메일을 받은 사람들
        /// </summary>
        /// <value>아웃룩 메일을 받은 사람들의 list</value>
        public List<MSGRecipientModel> Recipients
        {
            get { return this.recipients; }
        }
        private List<MSGRecipientModel> recipients = new List<MSGRecipientModel>();

        /// <summary>
        /// 아웃룩 메시지에 대한 첨부파일들
        /// </summary>
        /// <value>아웃룩 메시지에 첨부된 파일 list</value>
        public List<MSGAttachmentModel> Attachments
        {
            get { return this.attachments; }
        }
        private List<MSGAttachmentModel> attachments = new List<MSGAttachmentModel>();

        /// <summary>
        /// 아웃룩 메일 하위에 있는 메일들
        /// </summary>
        /// <value>이 아웃룩 메일 하위에 있는 메일에 대한 list</value>
        public List<MSGMessageModel> Messages
        {
            get { return this.messages; }
        }
        private List<MSGMessageModel> messages = new List<MSGMessageModel>();

        /// <summary>
        /// 누가 보냈는지에 대한 정보(이름)
        /// </summary>
        private String _from;
        public String From
        {
            get { return _from; }
        }

        /// <summary>
        /// 보낸 사람의 이메일 주소
        /// </summary>
        private String _fromAdd;
        public String FromAdd
        {
            get { return _fromAdd; }
        }

        /// <summary>
        /// 아웃룩 메일 제목
        /// </summary>
        private String _subject;
        public String Subject
        {
            get { return _subject; }
        }

        /// <summary>
        /// 아웃룩 메일 내용
        /// </summary>
        private String _bodyText;
        public String BodyText
        {
            get { return _bodyText; }
        }

        /// <summary>
        /// RTF로 암호화된 아웃룩 메일 내용
        /// </summary>
        public byte[] BodyByte { get; set; }
        private String _bodyRTF;
        public String BodyRTF
        {
            get
            {
                return _bodyRTF;
            }
        }

        /// <summary>
        /// 임시 폴더
        /// </summary>
        public string TempPath { get; set; }
        #endregion

        #region Constructor(s)

        /// <summary>
        /// Initializes a new instance of the <see cref="Message"/> class from a msg file.
        /// </summary>
        /// <param name="filename">The msg file to load.</param>
        public MSGMessageModel(string msgfile) : base(msgfile) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="Message"/> class from a <see cref="Stream"/> containing an IStorage.
        /// </summary>
        /// <param name="storageStream">The <see cref="Stream"/> containing an IStorage.</param>
        public MSGMessageModel(Stream storageStream) : base(storageStream) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="Message"/> class on the specified <see cref="NativeMethods.IStorage"/>.
        /// </summary>
        /// <param name="storage">The storage to create the <see cref="Message"/> on.</param>
        public MSGMessageModel(NativeCOMMethods.IStorage storage)
            : base(storage)
        {
            this.propHeaderSize = MSGFileEnv.PROPERTIES_STREAM_HEADER_TOP;
        }

        #endregion

        #region Methods(LoadStorage)

        /// <summary>
        /// 실제 데이터 매핑 순서는 대충 이렇다.
        /// 1. 깊이우선 탐색으로 stream과 property를 모두 꺼낸다
        /// 2. 꺼낸 stream과 property를 사용해 데이터를 변수에 넣는다
        /// Processes sub storages on the specified storage to capture attachment and recipient data.
        /// </summary>
        /// <param name="storage">The storage to check for attachment and recipient data.</param>
        public override void LoadStorage(NativeCOMMethods.IStorage storage)
        {
            base.LoadStorage(storage);

            foreach (ComTypes.STATSTG storageStat in this.subStorageStatistics.Values)
            {
                //element is a storage. get it and add its statistics object to the sub storage dictionary
                NativeCOMMethods.IStorage subStorage = this.storage.OpenStorage(storageStat.pwcsName, IntPtr.Zero, NativeCOMMethods.STGM.READ | NativeCOMMethods.STGM.SHARE_EXCLUSIVE, IntPtr.Zero, 0);

                //run specific load method depending on sub storage name prefix
                if (storageStat.pwcsName.StartsWith(MSGFileEnv.RECIP_STORAGE_PREFIX))
                {
                    MSGRecipientModel recipient = new MSGRecipientModel(new MSGFileReadBaseModel(subStorage));
                    this.recipients.Add(recipient);
                }
                else if (storageStat.pwcsName.StartsWith(MSGFileEnv.ATTACH_STORAGE_PREFIX))
                {
                    this.LoadAttachmentStorage(subStorage);
                }
                else
                {
                    //release sub storage
                    Marshal.ReleaseComObject(subStorage);
                }
            }

            this.SetMapiThisProperty();
        }

        /// <summary>
        /// Loads the attachment data out of the specified storage.
        /// </summary>
        /// <param name="storage">The attachment storage.</param>
        private void LoadAttachmentStorage(NativeCOMMethods.IStorage storage)
        {
            //create attachment from attachment storage
            MSGAttachmentModel attachment = new MSGAttachmentModel(new MSGFileReadBaseModel(storage));

            //if attachment is a embeded msg handle differently than an normal attachment
            string fileAttachData = MSGFileEnv.FileAttachmentTag.PR_ATTACH_METHOD.ToString("X").Substring(4);
            int attachMethod = attachment.GetMapiPropertyInt32(fileAttachData); // byte to int 의 문제
            if (attachMethod == MSGFileEnv.ATTACH_EMBEDDED_MSG)
            {
                //create new Message and set parent and header size
                MSGMessageModel subMsg = new MSGMessageModel(attachment.GetMapiProperty(fileAttachData) as NativeCOMMethods.IStorage);
                subMsg.parentMessage = this;
                subMsg.propHeaderSize = MSGFileEnv.PROPERTIES_STREAM_HEADER_EMBEDED;

                //add to messages list
                this.messages.Add(subMsg);
            }
            else
            {
                //add attachment to attachment list
                this.attachments.Add(attachment);
            }
        }

        /// <summary>
        /// 실제 변수에 데이터를 넣어줌
        /// </summary>
        public override void SetMapiThisProperty()
        {
            MSGMessageModel setModel = this;

            foreach (var stream in streamStatistics)
            {
                if (stream.Key.StartsWith(MSGFileEnv.DATA_STORAGE)) // stream에 데이터가 있는 경우 여기서 다 들어감
                {
                    //FileMessageTag propTag = (FileMessageTag)Int32.Parse(stream.Key.Substring(12, 8), System.Globalization.NumberStyles.HexNumber);
                    string tagnum = stream.Key.Substring(12, 4);
                    MSGFileEnv.FileMessageTag propTag = (MSGFileEnv.FileMessageTag)Int32.Parse(tagnum, System.Globalization.NumberStyles.HexNumber);
                    NativeCOMMethods.OutLookMAPI propType = (NativeCOMMethods.OutLookMAPI)ushort.Parse(stream.Key.Substring(16, 4), System.Globalization.NumberStyles.HexNumber);

                    byte[] getData = GetStreamBytes(stream.Value);

                    switch (propTag)
                    {
                        case MSGFileEnv.FileMessageTag.PR_SUBJECT:
                            this._subject = ByteToString(getData, propType);
                            break;

                        case MSGFileEnv.FileMessageTag.PR_BODY:
                            this._bodyText = ByteToString(getData, propType);
                            break;

                        case MSGFileEnv.FileMessageTag.PR_RTF_COMPRESSED:
                            if (getData == null || getData.Length == 0)
                            {
                                this._bodyRTF = null;
                            }
                            else
                            {
                                getData = CLZF.decompressRTF(getData);
                                this.BodyByte = getData;

                                File.WriteAllBytes(@"C:\Users\KIM CHAE YEON\Downloads\extracted.rtf", getData);
                                this._bodyRTF = Encoding.ASCII.GetString(getData);
                            }

                            break;

                        case MSGFileEnv.FileMessageTag.PR_SENDER_NAME:
                            this._from = ByteToString(getData, propType);
                            break;

                        case MSGFileEnv.FileMessageTag.PR_SENDER_EMAIL:
                            this._fromAdd = ByteToString(getData, propType);
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


        #endregion

        #region Methods(Disposing)

        protected override void Disposing()
        {
            //dispose sub storages
            foreach (MSGFileReadBaseModel subMsg in this.messages)
            {
                subMsg.Dispose();
            }

            //dispose sub storages
            foreach (MSGFileReadBaseModel recip in this.recipients)
            {
                recip.Dispose();
            }

            //dispose sub storages
            foreach (MSGFileReadBaseModel attach in this.attachments)
            {
                attach.Dispose();
            }
        }

        #endregion
    }
}

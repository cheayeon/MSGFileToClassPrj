using MSGFileToClassPrj.Enviroment;
using System;
using static MSGFileToClassPrj.Enviroment.MSGFileEnv;

namespace MSGFileToClassPrj.Models
{
    public class MSGRecipientModel : MSGFileReadBaseModel
    {
        #region Property(s)

        private string _displayName;
        /// <summary>
        /// 수신자의 이름
        /// </summary>
        public string DisplayName
        {
            get
            {
                return _displayName;
            }
        }

        private string _email;
        /// <summary>
        /// 수신자의 이메일
        /// </summary>
        public string Email
        {
            get
            {

                return _email;
            }
        }

        private RecipientType _type;
        /// <summary>
        /// 수신자 타입
        /// </summary>
        public RecipientType Type
        {
            get
            {
                if(_type == RecipientType.UnDefine)
                {
                    int getType = 0;
                    try
                    {
                        getType = (int)GetMapiPropertyFromPropertyByte(MSGFileEnv.FileRecipientTag.PR_RECIPIENT_TYPE.ToString("X").Substring(4));
                    }
                    catch (Exception) {
                        getType = 3;
                    }

                    _type = (RecipientType)getType;
                }

                return _type;
            }
        }

        #endregion

        public MSGRecipientModel(MSGFileReadBaseModel msgFileReadBaseModel)
                : base(msgFileReadBaseModel.storage)
        {
            this.propHeaderSize = MSGFileEnv.PROPERTIES_STREAM_HEADER_ATTACH_OR_RECIP;
            this.SetMapiThisProperty();
        }

        public override void SetMapiThisProperty()
        {
            MSGRecipientModel setModel = this;

            foreach (var stream in streamStatistics)
            {
                if (stream.Key.StartsWith(MSGFileEnv.DATA_STORAGE))
                {
                    string tagnum = stream.Key.Substring(12, 4);
                    MSGFileEnv.FileRecipientTag propTag = (MSGFileEnv.FileRecipientTag)Int32.Parse(tagnum, System.Globalization.NumberStyles.HexNumber);
                    NativeCOMMethods.OutLookMAPI propType = (NativeCOMMethods.OutLookMAPI)ushort.Parse(stream.Key.Substring(16, 4), System.Globalization.NumberStyles.HexNumber);

                    byte[] getData = GetStreamBytes(stream.Value);

                    switch (propTag)
                    {
                        case MSGFileEnv.FileRecipientTag.PR_EMAIL:
                        case MSGFileEnv.FileRecipientTag.PR_EMAIL_2:
                        case MSGFileEnv.FileRecipientTag.PR_EMAIL_3:
                            this._email = ByteToString(getData, propType);
                            break;
                        case MSGFileEnv.FileRecipientTag.PR_DISPLAY_NAME:
                        case MSGFileEnv.FileRecipientTag.PR_DISPLAY_CC:
                            this._displayName = ByteToString(getData, propType);
                            break;
                        case MSGFileEnv.FileRecipientTag.PR_RECIPIENT_TYPE:
                            this._type = (MSGFileEnv.RecipientType)GetMapiPropertyInt32(tagnum);
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
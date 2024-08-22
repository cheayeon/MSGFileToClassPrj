using MSGFileToClassPrj.Enviroment;
using System;
using ComTypes = System.Runtime.InteropServices.ComTypes;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MSGFileToClassPrj.Enviroment.Mannager;
using System.Runtime.InteropServices;
using System.IO;

namespace MSGFileToClassPrj.Models
{
    public class MSGFileReadBaseModel : IDisposable
    {
        /// <summary>
        /// 헤더 Byte Size (어떤 환경변수인지에 따라 헤더값이 달라짐)
        /// </summary>
        public int propHeaderSize = MSGFileEnv.PROPERTIES_STREAM_HEADER_TOP;

        #region 처음 만들어질 때 이후 사용 불가능 변수들
        /// 계속 사용하기 위해서는 파일을 계속 읽을 필요가 있음
        /// 즉, 통로를 닫지 않고 계속 열어둬야한다는 뜻임
        /// 언제 쓸지도 모르는 상태로 열어두는 것 보다는 데이터를 저장한 다음, 닫아두는게 더 나음.
        /// 이걸 계속 쓸라면 파일을 계속 열고 닫고 해줘야하는데, 이러면 시간이 좀 오래 걸릴 것 같음.
        /// 파일에서 원하는 데이터를 가져오면 가져올 수록 파일 전체를 가져와 메모리에 올리는 것과 같아짐.
        /// 어차피 파일의 데이터를 모두 올린다면 결국 속도 싸움이 됨.
        /// 하드드라이브에 접근하는 것 보다는 메모리에 접근하는 것이 속도 방면에서 유리함.
        /// 그래서 한번에 올린 후, 원하는 데이터에 접근하도록 수정함.

        /// <summary>
        /// 내 Stroage = 데이터 모음집(?)
        /// </summary>
        public NativeCOMMethods.IStorage storage;

        /// <summary>
        /// 이 Storage의 안에 있는 세부 Stream 의 모음
        /// </summary>
        public Dictionary<string, ComTypes.STATSTG> streamStatistics = new Dictionary<string, ComTypes.STATSTG>();

        /// <summary>
        /// 이 Storage의 안에 있는 세부 Storage의 모음
        /// </summary>
        public Dictionary<string, ComTypes.STATSTG> subStorageStatistics = new Dictionary<string, ComTypes.STATSTG>();
        #endregion

        /// <summary>
        /// 각 Key와 data들의 값
        /// </summary>
        public Dictionary<string, byte[]> datasDictionary = new Dictionary<string, byte[]>();
        public Dictionary<string, NativeCOMMethods.OutLookMAPI> encodingTypeDictionary = new Dictionary<string, NativeCOMMethods.OutLookMAPI>();

        /// <summary>
        /// 읽은 파일의 strorage 와 stream을 분리해 데이터를 뽑아 낼 수 있는 stream을 추출해야 한다.
        /// </summary>
        /// <param name="storage">하위 storage와 stream을 뽑아낼 수 있는 중추</param>
        public virtual void LoadStorage(NativeCOMMethods.IStorage storage)
        {
            this.storage = storage;

            //ensures memory is released
            ReferenceManager.AddItem(this.storage);

            NativeCOMMethods.IEnumSTATSTG storageElementEnum = null;
            try
            {
                //enum all elements of the storage
                storage.EnumElements(0, IntPtr.Zero, 0, out storageElementEnum);

                //iterate elements
                while (true)
                {
                    //get 1 element out of the com enumerator
                    uint elementStatCount;
                    ComTypes.STATSTG[] elementStats = new ComTypes.STATSTG[1]; // 뜻 = ComTypes.STATSTG 자료형을 가진 배열크기 1인 배열을 만들겠다.
                    storageElementEnum.Next(1, elementStats, out elementStatCount);

                    //break loop if element not retrieved
                    if (elementStatCount != 1)
                    {
                        break;
                    }

                    ComTypes.STATSTG elementStat = elementStats[0];
                    // https://learn.microsoft.com/ko-kr/windows/win32/api/objidl/ne-objidl-stgty
                    // 위의 경로에 정의된 값을 통해 가져온 element 값이 stroage 인지 stream인지 판단 가능
                    switch (elementStat.type)
                    {
                        case 1:
                            //element is a storage. add its statistics object to the storage dictionary
                            subStorageStatistics.Add(elementStat.pwcsName, elementStat);
                            break;

                        case 2:
                            //element is a stream. add its statistics object to the stream dictionary
                            streamStatistics.Add(elementStat.pwcsName, elementStat);
                            break;
                    }
                }
            }
            finally
            {
                //free memory
                if (storageElementEnum != null)
                {
                    Marshal.ReleaseComObject(storageElementEnum);
                }
            }
        }

        #region 파일의 값을 변수에 넣어줄때 호출
        /// <summary>
        /// 데이터를 실제 변수에 넣어줄때 호출하길....
        /// </summary>
        public virtual void SetMapiThisProperty()
        {

        }
        #endregion

        #region BaseFileStructClass 생성자
        /// <summary>
        /// 파일 경로를 사용해 새로운 <see cref="MSGFileReadBaseModel"/>클래스를 생성한다.
        /// </summary>
        /// <param name="storageFilePath">파일이 로드될 경로.</param>
        public MSGFileReadBaseModel(string storageFilePath)
        {
            //ensure provided file is an IStorage
            //파일이 IStorage 형식인지 확인
            if (NativeCOMMethods.StgIsStorageFile(storageFilePath) != 0)
            {
                throw new ArgumentException("The provided file is not a valid IStorage", "storageFilePath");
            }

            //open and load IStorage from file
            //IStorage를 사용해 파일을 연다.
            NativeCOMMethods.IStorage fileStorage;
            NativeCOMMethods.StgOpenStorage(storageFilePath, null, NativeCOMMethods.STGM.READ | NativeCOMMethods.STGM.SHARE_DENY_WRITE, IntPtr.Zero, 0, out fileStorage);
            this.LoadStorage(fileStorage);
        }


        /// <summary>
        /// 새로운 <see cref="MSGFileReadBaseModel"/>클래스를 매개변수 <see cref="Stream"/>을 IStorage로 감싸 만든다.
        /// </summary>
        /// <param name="storageStream">IStorage를 포함한 매개변수 <see cref="Stream"/>.</param>
        public MSGFileReadBaseModel(Stream storageStream)
        {
            if(storageStream == null)
            {
                throw new ArgumentNullException("StreamNull", "MSG File Stream is null");
            }

            NativeCOMMethods.IStorage memoryStorage = null;
            NativeCOMMethods.ILockBytes memoryStorageBytes = null;
            try
            {
                //read stream into buffer
                //stream을 버퍼를 통해 읽어온다.
                byte[] buffer = new byte[storageStream.Length];
                storageStream.Read(buffer, 0, buffer.Length);

                //create a ILockBytes (unmanaged byte array) and write buffer into it
                //ILockBytes 형식의 배열을 만든 후 버퍼의 내용을 쓴다.
                NativeCOMMethods.CreateILockBytesOnHGlobal(IntPtr.Zero, true, out memoryStorageBytes);
                memoryStorageBytes.WriteAt(0, buffer, buffer.Length, null);

                //ensure provided stream data is an IStorage
                //IStorage 형식으로 stream 데이터가 제공되는지 확인한다.
                if (NativeCOMMethods.StgIsStorageILockBytes(memoryStorageBytes) != 0)
                {
                    throw new ArgumentException("The provided stream is not a valid IStorage", "storageStream");
                }

                //open and load IStorage on the ILockBytes
                //ILockBytes배열에서 IStorage형식으로 데이터를 받아온다.
                NativeCOMMethods.StgOpenStorageOnILockBytes(memoryStorageBytes, null, NativeCOMMethods.STGM.READ | NativeCOMMethods.STGM.SHARE_DENY_WRITE, IntPtr.Zero, 0, out memoryStorage);

                //성공적으로 데이터를 불러왔을때 아래의 함수를 실행
                this.LoadStorage(memoryStorage);
            }
            catch
            {
                if (memoryStorage != null)
                {
                    Marshal.ReleaseComObject(memoryStorage);
                }
            }
            finally
            {
                if (memoryStorageBytes != null)
                {
                    Marshal.ReleaseComObject(memoryStorageBytes);
                }
            }
        }

        /// <summary>
        /// 새로운 <see cref="MSGFileReadBaseModel"/>클래스를 매개변수 <see cref="NativeCOMMethods.IStorage"/>를 사용해 만든다.
        /// </summary>
        /// <param name="storage"><see cref="MSGFileReadBaseModel"/>를 만들기 위한 매개변수</param>
        public MSGFileReadBaseModel(NativeCOMMethods.IStorage storage)
        {
            this.LoadStorage(storage);
        }

        /// <summary>
        /// 가비지 콜렉션에 의해 이 함수가 회수되기전에
        /// <see cref="MSGFileReadBaseModel"/>를 정리
        /// </summary>
        ~MSGFileReadBaseModel()
        {
            this.Dispose();
        }

        #endregion

        #region Methods(ByteToString)
        /// <summary>
        /// Byte로 구해진 값을 string 값으로 바꿔주는 코드
        /// </summary>
        /// <param name="streamByte">Byte 값</param>
        /// <param name="mapiType">Byte의 인코딩된 정보</param>
        /// <returns>Byte를 String으로 바꾼 값</returns>
        public string ByteToString(object streamByte, NativeCOMMethods.OutLookMAPI mapiType)
        {
            StreamReader streamReader = null;
            switch (mapiType)
            {
                case NativeCOMMethods.OutLookMAPI.PT_STRING8:
                    streamReader = new StreamReader(new MemoryStream(streamByte as byte[]), Encoding.UTF8);
                    break;
                case NativeCOMMethods.OutLookMAPI.PT_UNICODE:
                    streamReader = new StreamReader(new MemoryStream(streamByte as byte[]), Encoding.Unicode);
                    break;
            }

            if (streamReader == null) return null;

            string streamContent = streamReader.ReadToEnd();
            streamReader.Close();

            return streamContent;
        }
        #endregion

        // 파일 데이터 위치를 사용해 원하는 데이터를 가져오기
        #region Search & Get Methods(GetStreamBytes, GetStreamAsString)
        /// <summary>
        /// 파일의 데이터 위치를 사용해 값을 가져오기
        /// </summary>
        /// <param name="streamStatStg">파일 데이터의 실제 주소정보</param>
        /// <returns>해당 주소에 접근해 얻은 실제 데이터값</returns>
        public byte[] GetStreamBytes(ComTypes.STATSTG streamStatStg)
        {
            byte[] iStreamContent;
            ComTypes.IStream stream = null;
            try
            {
                //open stream from the storage
                stream = this.storage.OpenStream(streamStatStg.pwcsName, IntPtr.Zero, NativeCOMMethods.STGM.READ | NativeCOMMethods.STGM.SHARE_EXCLUSIVE, 0);

                //read the stream into a managed byte array
                iStreamContent = new byte[streamStatStg.cbSize];
                stream.Read(iStreamContent, iStreamContent.Length, IntPtr.Zero);
            }
            finally
            {
                if (stream != null)
                {
                    Marshal.ReleaseComObject(stream);
                }
            }

            //return the stream bytes
            return iStreamContent;
        }
        #endregion

        // 환경변수를 사용해 원하는 데이터를 가져오기(사용 안하는 중인듯? 다 만들고 나서도 안쓰면 삭제해도 될지도....)
        #region Methods(GetMapiProperty)

        /// <summary>
        /// property stream에 저장된 값을 가져옴
        /// </summary>
        /// <param name="propIdentifier">MSG 환경변수의 헥사 코드 string (4Byte)</param>
        /// <returns>프로퍼티에서 찾은 값 / 없으면 null</returns>
        public object GetMapiPropertyFromPropertyByte(string propIdentifier)
        {
            //propertys 가 dictionary에 없으면 null
            if (!this.datasDictionary.ContainsKey(MSGFileEnv.PROPERTIES_STREAM))
            {
                return null;
            }

            //property stream에서 raw 데이터 가져오기
            byte[] propBytes = datasDictionary[MSGFileEnv.PROPERTIES_STREAM];

            // 헤더를 제외한 내용을 16바이트 단위로 호출
            for (int i = this.propHeaderSize; i < propBytes.Length; i = i + 16)
            {
                // 해당 값이 어떤 자료형을 가졌는지 가져온다.
                ushort propType = BitConverter.ToUInt16(propBytes, i);

                // 값에 대한 환경변수 Hex코드를 가져온다.
                byte[] propIdent = new byte[] { propBytes[i + 3], propBytes[i + 2] };
                string propIdentString = BitConverter.ToString(propIdent).Replace("-", "");

                // 찾고있는 데이터가 아니면 다음것을 확인
                if (propIdentString != propIdentifier)
                {
                    continue;
                }

                // 데이터만 가져오기 위해서 값에대한 기본 정보를 제외
                switch (propType)
                {
                    case 2: // 16 bit int
                        return BitConverter.ToInt16(propBytes, i + 8);

                    case 3: // 32 bit int
                        return BitConverter.ToInt32(propBytes, i + 8);

                    case 64: // 64 bit int
                        return BitConverter.ToInt64(propBytes, i + 8);

                    default:
                        throw new ApplicationException("MAPI property has an unsupported type and can not be retrieved.");
                }
            }

            // 프로퍼티에 해당 내용이 없으면 null 반환
            return null;
        }

        /// <summary>
        /// storage 와 stream에 저장되었던 데이터 값을 가져옴
        /// </summary>
        /// <param name="propIdentifier">MSG 환경변수의 헥사 코드 string (4Byte)</param>
        /// <returns>매핑된 파일 데이터 값 또는 null</returns>
        private object GetMapiPropertyFromStreamOrStorage(string propIdentifier)
        {
            //get list of stream and storage identifiers which map to properties
            List<string> propKeys = new List<string>();
            propKeys.AddRange(this.datasDictionary.Keys);

            //determine if the property identifier is in a stream or sub storage
            string propTag = null;
            NativeCOMMethods.OutLookMAPI propType = NativeCOMMethods.OutLookMAPI.PT_UNSPECIFIED;
            byte[] propByte = null;
            foreach (string propKey in propKeys)
            {
                if (propKey.StartsWith(MSGFileEnv.DATA_STORAGE + propIdentifier))
                {
                    propTag = propKey.Substring(12, 8);
                    propType = (NativeCOMMethods.OutLookMAPI)ushort.Parse(propKey.Substring(16, 4), System.Globalization.NumberStyles.HexNumber);

                    propByte = datasDictionary[propKey];
                    break;
                }
            }

            //depending on prop type use method to get property value
            string containerName = MSGFileEnv.DATA_STORAGE + propTag;
            switch (propType)
            {
                case NativeCOMMethods.OutLookMAPI.PT_UNSPECIFIED:
                    return null;

                case NativeCOMMethods.OutLookMAPI.PT_STRING8:
                    return this.ByteToString(propByte, propType);

                case NativeCOMMethods.OutLookMAPI.PT_UNICODE:
                    return this.ByteToString(propByte, propType);

                case NativeCOMMethods.OutLookMAPI.PT_BINARY:
                    return propByte;
                    
                default:
                    throw new ApplicationException("MAPI property has an unsupported type and can not be retrieved.");
            }
        }

        /// <summary>
        /// 전달 받은 환경변수 값으로 실제 파일의 데이터 값을 가져옴
        /// </summary>
        /// <param name="propIdentifier">MSG 환경변수의 헥사 코드 string</param>
        /// <returns>실제 파일 데이터값(raw)</returns>
        public object GetMapiProperty(string propIdentifier)
        {
            //try get prop value from stream or storage
            //stream 또는 storage에 정의된 환경변수라면 아래의 시도로 값을 가져올 수 있다.
            object propValue = this.GetMapiPropertyFromStreamOrStorage(propIdentifier);

            //if not found in stream or storage try get prop value from property stream
            //stream과 storage에 정의된 값이 아니라면 property stream에서 값을 가져올수있다.
            if (propValue == null)
            {
                propValue = this.GetMapiPropertyFromPropertyByte(propIdentifier);
            }

            return propValue;
        }

        /// <summary>
        /// 파일에서 전달해준 환경변수에 정의된 값을 string로 전달
        /// </summary>
        /// <param name="propIdentifier">MSG 환경변수의 헥사 코드 string</param>
        /// <returns>실제 값을 string으로 전달</returns>
        public string GetMapiPropertyString(string propIdentifier)
        {
            return this.GetMapiProperty(propIdentifier) as string;
        }

        /// <summary>
        /// 파일에서 전달해준 환경변수에 정의된 값을 Int16로 전달
        /// </summary>
        /// <param name="propIdentifier">MSG 환경변수의 헥사 코드 string</param>
        /// <returns>The value of the MAPI property as a short.</returns>
        public Int16 GetMapiPropertyInt16(string propIdentifier)
        {
            return (Int16)this.GetMapiProperty(propIdentifier);
        }

        /// <summary>
        /// 파일에서 전달해준 환경변수에 정의된 값을 int로 전달
        /// </summary>
        /// <param name="propIdentifier">MSG 환경변수의 헥사 코드 string</param>
        /// <returns>The value of the MAPI property as a integer.</returns>
        public int GetMapiPropertyInt32(string propIdentifier)
        {
            return (int)this.GetMapiProperty(propIdentifier);
        }

        /// <summary>
        /// 파일에서 전달해준 환경변수에 정의된 값을 byte[]로 전달
        /// </summary>
        /// <param name="propIdentifier">MSG 환경변수의 헥사 코드 string</param>
        /// <returns>The value of the MAPI property as a byte array.</returns>
        public byte[] GetMapiPropertyBytes(string propIdentifier)
        {
            return (byte[])this.GetMapiProperty(propIdentifier);
        }

        /// <summary>
        /// 파일에서 전달해준 환경변수에 정의된 값을 object로 전달
        /// </summary>
        /// <param name="propIdentifier">MSG 환경변수의 헥사 코드 string</param>
        /// <returns></returns>
        public object GetMapiPropertyObject(string propIdentifier)
        {
            return this.GetMapiProperty(propIdentifier);
        }

        #endregion

        #region IDisposable Members
        /// <summary>
        /// 이 객체의 dispose(여기서는 파일열기유지여부) 상태
        /// </summary>
        private bool disposed = false;

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// 관리되지 않는 리소스 해제, 해제 또는 재설정과 관련된 애플리케이션 정의 작업을 수행합니다.
        /// </summary>
        public void Dispose()
        {
            if (!this.disposed)
            {
                //ensure only disposed once
                //한번만 폐기
                this.disposed = true;

                //call virtual disposing method to let sub classes clean up
                //하위 클래스의 폐기 진행을 위해 가상 메소드 호출
                this.Disposing();

                //release COM storage object and suppress finalizer
                //COM storage 객체를 해제하고 finalizer 호출을 막는다.
                //finalizer는 GC(가비지콜렉터)가 더이상 사용하지 않는 객체를 정리할때 사용하는 메소드다.
                if (this.storage != null)
                {
                    ReferenceManager.RemoveItem(this.storage);
                    Marshal.ReleaseComObject(this.storage); // 객체 해제 = 사용안함 = 메모리 free
                    GC.SuppressFinalize(this);              // finalizer 억제
                }
            }
        }

        /// <summary>
        /// Gives sub classes the chance to free resources during object disposal.
        /// </summary>
        protected virtual void Disposing() { }

        #endregion
    }
}

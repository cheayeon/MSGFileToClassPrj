using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MSGFileToClassPrj.Enviroment
{
    /// <summary>
    /// https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxprops/f6ab1613-aefe-447d-a49c-18217230b148
    /// MSG 파일과 관련하여 더 필요한 데이터가 있다면 다음과 같은 순서로 값을 추가할 수 있다.
    /// 1. 위의 경로에서 최신 환경변수 정의 파일을 다운받는다.
    /// 2. 원하는 데이터를 파일에서 찾는다.
    /// 3. Alternate names로 정의된 이름과 Property ID로 정의된 데이터값을 사용해 아래의 방식대로 값을 추가한다.
    /// 4. 파일을 불러올 때 원하는 변수에다가 데이터를 추가할 수 있도록 설정한다.
    /// </summary>

    public class MSGFileEnv
    {
        // 첨부파일 데이터들
        public enum FileAttachmentTag : Int32
        {
            PR_ATTACH_FILENAME = 0x3704,        // 첨부파일 이름
            PR_ATTACH_LONG_FILENAME = 0x3707,   // 첨부파일 긴 이름
            PR_ATTACH_DATA = 0x3701,            // 첨부파일 데이터
            PR_ATTACH_METHOD = 0x3705,          // 첨부파일에 엑세스하는 방식
            PR_ATTACHMENT_LINKID = 0x7FFA,      // 첨부파일에 링크된 객체 유형
            PR_RENDERING_POSITION = 0x370B,     // 첨부파일의 렌더링 제어 (파일이 열리고 난 후에 데이터들의 위치를 정의하는듯. (ex/ 위아래 여백))
            PR_ATTACH_CONTENT_ID = 0x3712,
            PR_ATTACHMENT_CONTACTPHOTO = 0x7FFF
        }

        // 수신자, 참조자 데이터들
        public enum FileRecipientTag : Int32
        {
            PR_DISPLAY_NAME = 0x3001,   // 수신자 이름
            PR_DISPLAY_CC = 0x0E03,     // 참조자 이름
            PR_EMAIL = 0x39FE,          // 이메일 ver1
            PR_EMAIL_2 = 0x403E,        // 이메일 ver2
            PR_EMAIL_3 = 0x3003,        // 이메일 ver3
            PR_RECIPIENT_TYPE = 0x0C15, // 수신자(1) or 참조(2)
        }

        // 실제 메시지 데이터들
        public enum FileMessageTag : Int32
        {
            PR_SUBJECT = 0x0037,        // 제목
            PR_BODY = 0x1000,           // 내용
            PR_RTF_COMPRESSED = 0x1009, // RTF로 암호화된 내용
            PR_SENDER_NAME = 0x0C1A,    // 발신자 이름
            PR_SENDER_EMAIL = 0x0C1F,   // 발신자 메일 주소
            PR_ORIGINAL_DELIVERY_TIME = 0x0E06,   // 메시지 전송시간 
        }

        public enum RecipientType
        {
            UnDefine = 0,
            To = 1,
            CC = 2,
            Unknown = 3
        }

        //attachment constants
        // 첨부파일 이름 및 데이터들
        public const string ATTACH_STORAGE_PREFIX = "__attach_version1.0_#";
        public const int ATTACH_BY_VALUE = 1;
        public const int ATTACH_EMBEDDED_MSG = 5;
        public const int ATTACH_OLE = 6;

        //recipient constants
        // 수신자 및 참조자 정보
        public const string RECIP_STORAGE_PREFIX = "__recip_version1.0_#";
        public const int MAPI_TO = 1;
        public const int MAPI_CC = 2;

        //property stream constants
        // 각 stream이 포함하는 내용들
        public const string PROPERTIES_STREAM = "__properties_version1.0";

        // 각 실제 데이터의 헤더길이
        // 각 숫자의 기준은 byte 수다.
        // stream 단위로 데이터를 받아올때 사용해야한다.
        public const int PROPERTIES_STREAM_HEADER_TOP = 32;    // 예약(8) 수신자 데이터 저장소 ID(4) 첨부파일 데이터 저장소 ID(4) 수신자 수(4) 첨부파일 수(4) 예약(8)
                                                                // 위의 구조에서 예약은 무시한다. (예약 = 쓰기 전용 바이트)
        public const int PROPERTIES_STREAM_HEADER_EMBEDED = 24;    // 예약(8) 다음 수신자 데이터 저장소 ID(4) 다음 첨부파일 데이터 저장소 ID(4) 수신자 수 (4) 첨부파일 수(4)
                                                                    // 위의 구조에서 예약은 무시한다. (예약 = 쓰기 전용 바이트)
        public const int PROPERTIES_STREAM_HEADER_ATTACH_OR_RECIP = 8; // 예약만 무시하면 됨. (예약 = 쓰기 전용 바이트)

        //name id storage name in root storage
        public const string NAMEID_STORAGE = "__nameid_version1.0";

        public const string DATA_STORAGE = "__substg1.0_";
    }
}

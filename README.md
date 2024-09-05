                // 파일에 아래의 특수문자가 들어가면 오류남.
                var pattern = "[|^\\?\"<>:]+";
                message.TempPath = MSGTempPath + "\\" + Regex.Replace(message.Subject, pattern,"") + ".html";

                추가할것 언젠가

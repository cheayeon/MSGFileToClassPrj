using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace MSGFileToClassPrj.Enviroment.Mannager
{
    // Cpp를 통해 사용한 메모리를 free 시켜주기 위한 manager
    public class ReferenceManager
    {
        public static void AddItem(object track)
        {
            lock (instance)
            {
                if (!instance.trackingObjects.Contains(track))
                {
                    instance.trackingObjects.Add(track);
                }
            }
        }

        public static void RemoveItem(object track)
        {
            lock (instance)
            {
                if (instance.trackingObjects.Contains(track))
                {
                    instance.trackingObjects.Remove(track);
                }
            }
        }

        private static ReferenceManager instance = new ReferenceManager();

        private List<object> trackingObjects = new List<object>();

        private ReferenceManager() { }

        ~ReferenceManager()
        {
            foreach (object trackingObject in trackingObjects)
            {
                Marshal.ReleaseComObject(trackingObject);
            }
        }
    }
}

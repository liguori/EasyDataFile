using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyDataFile.Import.ETLProvider
{
    public interface ISingleModelProviderETL
    {
        Type Model { get; set; }

        void Import(string fileName);

        Action<int, int, Object, Exception> RecordReadyMethod { get; set; }
    }
}

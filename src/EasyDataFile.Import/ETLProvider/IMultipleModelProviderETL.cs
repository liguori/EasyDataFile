using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SmartETL.ETLProvider
{
    public interface IMultipleModelProviderETL
    {
        List<Type> Models { get; set; }

        void AddModel(Type t);

        void Import(string fileName);

        Action<long, long, IEnumerable<Object>, Exception> RecordReadyMethod { get; set; }

        Action<object> ObecjtCreated { get; set; }
    }
}

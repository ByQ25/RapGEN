using System.Collections.Generic;

namespace RapGEN
{
    interface IDataLoader<T> : IProgressInfo
    {
        Queue<T> LoadData();
    }
}

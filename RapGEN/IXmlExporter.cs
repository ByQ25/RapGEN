using System.Collections.Generic;

namespace RapGEN
{
    interface IXmlExporter<T> : IProgressInfo
    {
        void Export(Queue<T> Data);
        void DeleteCreatedFile();
    }
}

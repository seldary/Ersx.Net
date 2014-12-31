using System;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace Ersx.Net
{
    public interface IResxSorter
    {
        XDocument Sort(XDocument resxDocument);
    }

    public class ResxSorter : IResxSorter
    {
        public XDocument Sort(XDocument resxDocument)
        {
            Func<XElement, string> name = _ => (string)_.Attribute("name");
            return new XDocument(
                new XElement(resxDocument.Root.Name,
                    resxDocument.Root.Nodes().Where(comment => comment.NodeType == XmlNodeType.Comment),
                    resxDocument.Root.Elements().Where(_ => _.Name.LocalName == "schema"),
                    resxDocument.Root.Elements("resheader").OrderBy(name),
                    resxDocument.Root.Elements("assembly").OrderBy(name),
                    resxDocument.Root.Elements("metadata").OrderBy(name),
                    resxDocument.Root.Elements("data").OrderBy(name)));
        }
    }
}

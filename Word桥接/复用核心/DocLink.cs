namespace AnnotationModal;
public class DocLink
{
    public string LinkPath { get; set; }
    public string RelativePath { get; set; }
    public string LinkAddress { get; set; }
    public string LinkName { get; set; }
    public string Guid { get; set; }
    public string Descr { get; set; }
    public DocLinkMode LinkMode { get; set; } = DocLinkMode.Address;
}

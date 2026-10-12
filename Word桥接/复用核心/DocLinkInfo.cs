using AnnotationModal;
namespace AnnotationDal;
// 原六个字段及默认模式原样保留；仅去除编译器生成属性和混淆保护。
public class DocLinkInfo
{
    public string Address { get; set; }
    public string Name { get; set; }
    public string FilePath { get; set; }
    public bool IsRelative { get; set; }
    public string Descr { get; set; }
    public DocLinkMode LinkMode { get; set; } = DocLinkMode.Address;
}

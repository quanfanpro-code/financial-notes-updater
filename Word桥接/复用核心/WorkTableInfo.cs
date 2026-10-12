namespace AnnotationModal;

public class WorkTableInfo
{
	public string SavePath { get; set; }

	public string RelativePath { get; set; }

	public string LinkName { get; set; }

	public int RowResize { get; set; }

	public int ColResize { get; set; }

	public int autoInsert { get; set; }

	public WorkTableInfo()
	{
		autoInsert = 0;
	}
}

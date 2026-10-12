namespace AnnotationModal;

public class DalSetting
{
	public string ZeroFormat { get; set; }

	public bool ShowColor { get; set; }

	public DalSetting()
	{
		ZeroFormat = "-";
		ShowColor = false;
	}
}

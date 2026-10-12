using AnnotationModal;
using Newtonsoft.Json;

namespace ResourceFactory;

public static class CommonFunction
{
	public static bool IsLinkTable(string descr)
	{
		// 原方法只补长度检查，避免一字说明触发越界；连接格式保持不变。
		if (!string.IsNullOrEmpty(descr) && descr.Length >= 4 && descr.Substring(0, 2) == "<#" && descr.Substring(descr.Length - 2, 2) == "#>")
		{
			return true;
		}
		return false;
	}

	public static WorkTableInfo GetWorkTableInfo(string descr)
	{
		return JsonConvert.DeserializeObject<WorkTableInfo>(descr.Substring(2, descr.Length - 4));
	}
}

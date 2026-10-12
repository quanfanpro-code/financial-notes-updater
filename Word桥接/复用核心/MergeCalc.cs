using 附注工具.核心;

namespace AnnotationDal;

public class MergeCalc
{
    // 原 MergeCalc.cs 90～119行；只把XSSFSheet合并清单改为已读取的坐标清单。
    public Dictionary<int, List<int>> GetMergeDic(IReadOnlyList<MergeRegion> sh)
    {
        Dictionary<int, List<int>> dictionary = new Dictionary<int, List<int>>();
        int numMergedRegions = sh.Count;
        List<int> list = null;
        for (int i = 0; i < numMergedRegions; i++)
        {
            MergeRegion mergedRegion = sh[i];
            for (int j = mergedRegion.FirstRow; j <= mergedRegion.LastRow; j++)
            {
                if (dictionary.ContainsKey(j))
                {
                    list = dictionary[j];
                }
                else
                {
                    list = new List<int>();
                    dictionary.Add(j, list);
                }
                for (int k = mergedRegion.FirstColumn + 1; k <= mergedRegion.LastColumn; k++)
                {
                    list.Add(k);
                }
            }
        }
        return dictionary;
    }

    // 原 MergeCalc.cs 185～214行；保留有效列的绝对坐标，只修补默认零尾部。
    public Dictionary<int, int[]> GetMapArray(MergeRegion cellRange, Dictionary<int, List<int>> MergeDic)
    {
        Dictionary<int, int[]> dictionary = new Dictionary<int, int[]>();
        int num = 0;
        List<int> list = null;
        for (int i = cellRange.FirstRow; i <= cellRange.LastRow; i++)
        {
            if (!MergeDic.ContainsKey(i))
            {
                continue;
            }
            int[] array = new int[cellRange.LastColumn - cellRange.FirstColumn + 1];
            list = MergeDic[i];
            num = 0;
            for (int j = cellRange.FirstColumn; j <= cellRange.LastColumn; j++)
            {
                if (!list.Contains(j))
                {
                    array[num] = j;
                    num++;
                }
            }
            Array.Resize(ref array, num);
            dictionary.Add(i, array);
        }
        return dictionary;
    }
}

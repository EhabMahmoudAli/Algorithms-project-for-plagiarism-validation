using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using OfficeOpenXml;
using System.IO;
using ConsoleApp1;
using System.Diagnostics;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;


class Vertex
{
    public string FilePath { get; }
    public double Similarity1 { get; }
    public double Similarity2 { get; }

    public Vertex(string filePath)
    {
        FilePath = filePath;

    }
}


class Edge
{
    public Vertex Source { get; }
    public Vertex Destination { get; }
    public double Similarity1 { get; }
    public double Similarity2 { get; }
    public int LinesMatched { get; }

    public Edge(Vertex source, Vertex destination, double similarity1, double similarity2, int linesMatched)
    {
        Source = source;
        Destination = destination;
        Similarity1 = similarity1;
        Similarity2 = similarity2;
        LinesMatched = linesMatched;
    }
}


class MyGraph
{
    private List<Vertex> vertices;
    private List<Edge> edges;

    public MyGraph()
    {
        vertices = new List<Vertex>();
        edges = new List<Edge>();
    }

    public List<Edge> GetEdges()
    {
        return edges;
    }
    //modifed part
    public List<Vertex> Getvertices()
    {
        return vertices;
    }

    public void AddVertex(string filePath)
    {
        Vertex existingVertex = vertices.Find(v => v.FilePath == filePath);
        if (existingVertex == null)
        {
            Vertex vertex = new Vertex(filePath);
            vertices.Add(vertex);
        }
    }

    public void AddEdge(string sourceFilePath, string destinationFilePath, double similarity1, double similarity2, int linesMatched)
    {
        Vertex source = vertices.Find(v => v.FilePath == sourceFilePath);
        Vertex destination = vertices.Find(v => v.FilePath == destinationFilePath);

        if (source != null && destination != null)
        {
            Edge existingEdge = edges.Find(e => e.Source == source && e.Destination == destination);
            if (existingEdge == null)
            {
                Edge edge = new Edge(source, destination, similarity1, similarity2, linesMatched);
                edges.Add(edge);
            }
        }

    }


    public List<List<Vertex>> FindConnectedComponents()
    {
        List<List<Vertex>> connectedComponents = new List<List<Vertex>>();
        bool[] visited = new bool[vertices.Count];

        for (int i = 0; i < vertices.Count; i++)
        {
            if (!visited[i])
            {
                List<Vertex> component = new List<Vertex>();
                DFS(vertices[i], visited, component);
                component.Sort((v1, v2) => int.Parse(v1.FilePath).CompareTo(int.Parse(v2.FilePath)));
                connectedComponents.Add(component);
            }
        }

        return connectedComponents;
    }
    public List<List<Edge>> FindConnectedComponentsforedges()
    {
        List<List<Edge>> connectedComponents = new List<List<Edge>>();
        bool[] visited = new bool[vertices.Count];

        for (int i = 0; i < vertices.Count; i++)
        {
            if (!visited[i])
            {
                List<Edge> componentEdges = new List<Edge>();
                DFSforedges(vertices[i], visited, componentEdges);
                connectedComponents.Add(componentEdges);
            }
        }

        return connectedComponents;
    }

    private void DFS(Vertex vertex, bool[] visited, List<Vertex> component)
    {
        visited[vertices.IndexOf(vertex)] = true;
        component.Add(vertex);

        foreach (var edge in edges)
        {
            if (edge.Source == vertex && !visited[vertices.IndexOf(edge.Destination)])
            {
                DFS(edge.Destination, visited, component);
            }
            else if (edge.Destination == vertex && !visited[vertices.IndexOf(edge.Source)])
            {
                DFS(edge.Source, visited, component);
            }
        }


    }
    private void DFSforedges(Vertex vertex, bool[] visited, List<Edge> componentEdges)
    {
        visited[vertices.IndexOf(vertex)] = true;

        foreach (var edge in edges)
        {
            if (edge.Source == vertex && !visited[vertices.IndexOf(edge.Destination)])
            {
                componentEdges.Add(edge);
                DFSforedges(edge.Destination, visited, componentEdges);
            }
            else if (edge.Destination == vertex && !visited[vertices.IndexOf(edge.Source)])
            {
                componentEdges.Add(edge);
                DFSforedges(edge.Source, visited, componentEdges);
            }
        }
    }

    public static double CalculateAverageSimilarity(List<Vertex> component, List<Edge> edges)
    {
        double totalSimilarity = 0;
        int edgeCount = 0;

        foreach (var edge in edges)
        {
            if (component.Contains(edge.Source) && component.Contains(edge.Destination))
            {
                totalSimilarity += edge.Similarity1 + edge.Similarity2;
                edgeCount++;
            }
        }

        if (edgeCount == 0)
        {
            return 0;
        }

        double averageSimilarity = totalSimilarity / (edgeCount * 2);
        return Math.Round(averageSimilarity, 1);
    }
    public static double CalculateAverageSimilarityforedges(List<Edge> component, List<Edge> edges)
    {
        double totalSimilarity = 0;
        int edgeCount = 0;

        foreach (var edge in edges)
        {
            if (component.Contains(edge))
            {
                totalSimilarity += edge.Similarity1 + edge.Similarity2;
                edgeCount++;
            }
        }

        if (edgeCount == 0)
        {
            return 0;
        }

        double averageSimilarity = totalSimilarity / (edgeCount * 2);
        return Math.Round(averageSimilarity, 1);
    }

    public List<Tuple<string, string, double, double, int>> MST(List<Vertex> vertices, List<Edge> edges)
    {
        /*List < List<Vertex> >allsets= new List<List<Vertex>>();
        List<Vertex> unionedset = new List<Vertex>();*/
        // bool edited = false;
        List<Tuple<string, string, double, double, int>> mst = new List<Tuple<string, string, double, double, int>>();
        /*        List<HashSet<Vertex>> ListOfSets = new List<HashSet<Vertex>>();
        */
        Dictionary<Vertex, Vertex> setOfEachVer = new Dictionary<Vertex, Vertex>();
        foreach (Vertex vertex in vertices)
        {
            // HashSet<Vertex> verSet = new HashSet<Vertex>();
            // verSet.Add(vertex);
            // ListOfSets.Add(verSet);
            setOfEachVer[vertex] = vertex;
        }
        edges.Sort((x, y) =>
        {
            if (Math.Max(x.Similarity1, x.Similarity2) != Math.Max(y.Similarity1, y.Similarity2))
            {
                return Math.Max(y.Similarity1, y.Similarity2).CompareTo(Math.Max(x.Similarity1, x.Similarity2));

            }
            else
            {
                return y.LinesMatched.CompareTo(x.LinesMatched);

            }
        });
        foreach (Edge e in edges)
        {
            if (setOfEachVer[e.Source] != setOfEachVer[e.Destination])
            {
                mst.Add(new Tuple<string, string, double, double, int>(e.Source.FilePath, e.Destination.FilePath, e.Similarity1, e.Similarity2, e.LinesMatched));


                //khaly kol ely tab3 el des b el source
                Vertex set1 = setOfEachVer[e.Destination];
                Vertex set2 = setOfEachVer[e.Source];
                setOfEachVer[e.Destination] = setOfEachVer[e.Source];
                // Create a copy of the dictionary before iterating
                Dictionary<Vertex, Vertex> setCopy = new Dictionary<Vertex, Vertex>(setOfEachVer);
                foreach (Vertex key in setCopy.Keys)
                {
                    if (setCopy[key] == set1)
                    {
                        // setCopy[key] = set2;

                        setOfEachVer[key] = set2;
                    }
                }


            }
        }
        /* foreach (HashSet<Vertex> vertexSet in ListOfSets)
         {
             vertexSet.Union(Source,dis);

         }*/
        //OrderByDescending();

        return mst;
    }

}


class ExcelReader
{

    public static List<Tuple<string, string, double, double, int>> ReadMatchingPairsFromExcel(string filePath)
    {
        List<Tuple<string, string, double, double, int>> matchingPairs = new List<Tuple<string, string, double, double, int>>();
        using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            int rowCount = worksheet.Dimension.Rows;

            for (int row = 2; row <= rowCount; row++)
            {
                string file1 = ExtractPartFromValue(worksheet.Cells[row, 1].Value?.ToString());
                string file2 = ExtractPartFromValue(worksheet.Cells[row, 2].Value?.ToString());

                double similarity1 = ExtractSimilarity1(worksheet.Cells[row, 1].Value?.ToString(), worksheet.Cells[row, 2].Value?.ToString());
                double similarity2 = ExtractSimilarity2(worksheet.Cells[row, 1].Value?.ToString(), worksheet.Cells[row, 2].Value?.ToString());
                int linesMatched = ExtractLinesMatched(worksheet.Cells[row, 3].Value?.ToString());

                if (!string.IsNullOrEmpty(file1) && !string.IsNullOrEmpty(file2))
                {
                    matchingPairs.Add(new Tuple<string, string, double, double, int>(file1, file2, similarity1, similarity2, linesMatched));
                }
            }
        }
        return matchingPairs;
    }

    private static int ExtractLinesMatched(string value)
    {
        if (value == null)
        {
            return 0;
        }

        int linesMatched;
        if (int.TryParse(value, out linesMatched))
        {
            return linesMatched;
        }

        return 0;
    }

    private static string ExtractPartFromValue(string value)
    {
        if (value == null)
        {
            return null;
        }

        int lastSlashIndex = value.LastIndexOf('/');

        if (lastSlashIndex != -1)
        {
            string substring = value.Substring(0, lastSlashIndex);
            int lastNonNumericIndex = 0;

            for (int i = lastSlashIndex - 1; i >= 0; i--)
            {
                if (!char.IsDigit(substring[i]))
                {
                    lastNonNumericIndex = i;
                    break;
                }
            }

            if (char.IsDigit(substring[lastNonNumericIndex]))
            {
                return substring;
            }
            else
            {
                //lw end of the string
                return substring.Substring(lastNonNumericIndex + 1);
            }
        }

        return null;
    }
    private static double ExtractSimilarity1(string value1, string value2)
    {
        if (value1 == null || value2 == null)
        {
            return 0.0;
        }

        int startIndex1 = value1.LastIndexOf('(') + 1;
        int endIndex1 = value1.LastIndexOf('%');
        if (startIndex1 >= 0 && endIndex1 >= 0 && endIndex1 > startIndex1)
        {
            string similarityString1 = value1.Substring(startIndex1, endIndex1 - startIndex1);
            double similarity1;
            if (double.TryParse(similarityString1, out similarity1))
            {
                return similarity1;
            }
        }

        return 0.0;
    }

    private static double ExtractSimilarity2(string value1, string value2)
    {
        if (value1 == null || value2 == null)
        {
            return 0.0;
        }

        int startIndex2 = value2.LastIndexOf('(') + 1;
        int endIndex2 = value2.LastIndexOf('%');
        if (startIndex2 >= 0 && endIndex2 >= 0 && endIndex2 > startIndex2)
        {
            string similarityString2 = value2.Substring(startIndex2, endIndex2 - startIndex2);
            double similarity2;
            if (double.TryParse(similarityString2, out similarity2))
            {
                return similarity2;
            }
        }

        return 0.0;
    }
}


class ExcelWriter
{
    public static void WriteToExcel(List<List<Vertex>> connectedComponents, List<Edge> edges, string outputPath)
    {
        using (ExcelPackage excelPackage = new ExcelPackage())
        {
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Results");

            worksheet.Cells[1, 1].Value = "Component Index";
            worksheet.Cells[1, 2].Value = "Vertices";
            worksheet.Cells[1, 3].Value = "Average Similarity";
            worksheet.Cells[1, 4].Value = "Component Count";

            connectedComponents = connectedComponents.OrderByDescending(component => MyGraph.CalculateAverageSimilarity(component, edges)).ToList();

            for (int i = 0; i < connectedComponents.Count; i++)
            {
                List<Vertex> component = connectedComponents[i];
                double averageSimilarity = MyGraph.CalculateAverageSimilarity(component, edges);
                List<string> sortedVertices = component.Select(v => v.FilePath).OrderBy(v => int.Parse(v)).ToList(); // sorting
                string vertices = string.Join(", ", sortedVertices);
                int componentCount = component.Count;
                worksheet.Cells[i + 2, 1].Value = i + 1;
                worksheet.Cells[i + 2, 2].Value = vertices;
                worksheet.Cells[i + 2, 3].Value = averageSimilarity;
                worksheet.Cells[i + 2, 4].Value = componentCount;
            }

            FileInfo excelFile = new FileInfo(outputPath);
            excelPackage.SaveAs(excelFile);
        }
    }
    
    public static void WriteMSTToExcel(List<Edge> edges, List<List<Edge>> MSTconnectedComponentsE, string outputPath)
    {
        using (ExcelPackage excelPackage = new ExcelPackage())
        {
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Results");

            worksheet.Cells[1, 1].Value = "file1";
            worksheet.Cells[1, 2].Value = "file2";
            worksheet.Cells[1, 3].Value = "lines matched";
            //worksheet.Cells[1, 4].Value = "Component Count";
            MSTconnectedComponentsE = MSTconnectedComponentsE.OrderByDescending(component => MyGraph.CalculateAverageSimilarityforedges(component, edges)).ToList();
            int row = 2;
            for (int i = 0; i < MSTconnectedComponentsE.Count; i++)
            {
                List<Edge> component = MSTconnectedComponentsE[i];
                component.Sort((x, y) => y.LinesMatched.CompareTo(x.LinesMatched));

                for (int j = 0; j < component.Count; j++)
                {
                    //int componentCount = component.Count;
                    worksheet.Cells[j + row, 1].Value = component[j].Source.FilePath;
                    worksheet.Cells[j + row, 2].Value = component[j].Destination.FilePath;
                    worksheet.Cells[j + row, 3].Value = component[j].LinesMatched;
                    // worksheet.Cells[i + 2, 4].Value =;

                }

                row += component.Count;


            }

            FileInfo excelFile = new FileInfo(outputPath);
            excelPackage.SaveAs(excelFile);
        }
    }

}


class Program
{
    static void Main(string[] args)
    {
        string filePath = @"D:\Test Cases\Complete\Medium\2-Input.xlsx";
        List<Tuple<string, string, double, double, int>> matchingPairs = ExcelReader.ReadMatchingPairsFromExcel(filePath);
        MyGraph g = ConstructGraph(matchingPairs);
        Stopwatch stopwatch = Stopwatch.StartNew();


        List<List<Vertex>> connectedComponents = g.FindConnectedComponents();
        //Mst functions
        List<Tuple<string, string, double, double, int>> mst = g.MST(g.Getvertices(), g.GetEdges());
        MyGraph MstGraph = ConstructGraph(mst);
        List<List<Vertex>> MSTconnectedComponentsV = MstGraph.FindConnectedComponents();
        List<List<Edge>> MSTconnectedComponentsE = MstGraph.FindConnectedComponentsforedges();

        


        ExcelWriter.WriteToExcel(connectedComponents, g.GetEdges(), @"D:\Test Cases\Sample\output.xlsx");
        ExcelWriter.WriteMSTToExcel(MstGraph.GetEdges(), MSTconnectedComponentsE, @"D:\Test Cases\Sample\MST.xlsx");

        stopwatch.Stop();
        double elapsedTimeInSeconds = stopwatch.Elapsed.TotalSeconds;
        Console.WriteLine($"Execution time: {elapsedTimeInSeconds * 1000} MilliSeconds");


        Console.WriteLine("generating the files is done");
        Console.ReadKey();
    }

    static MyGraph ConstructGraph(List<Tuple<string, string, double, double, int>> matchingPairs)
    {
        MyGraph g = new MyGraph();
        foreach (var tuple in matchingPairs)
        {
            string file1 = tuple.Item1;
            string file2 = tuple.Item2;
            double similarity1 = tuple.Item3;
            double similarity2 = tuple.Item4;
            int linesMatched = tuple.Item5;

            g.AddVertex(file1);
            g.AddVertex(file2);
            g.AddEdge(file1, file2, similarity1, similarity2, linesMatched);
        }

        return g;
    }

}
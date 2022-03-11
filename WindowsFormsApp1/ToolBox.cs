using System;
using System.Collections;
using System.Collections.Concurrent;

public class Tool
{
    public Tool(int id, string name)
    {
        this.id = id;
        this.tName = name;
    }

    public int id;
    public bool status;
    public bool silent = false;
    public string tName;
    public string tPath;
    public string tType;
    public string buttonName;
    public string verify;
    public string level;
}

public class Tools : IEnumerable
{
    private Tool[] _tools;
    public Tools(Tool[] tArray)
    {
        _tools = new Tool[tArray.Length];
        for (int i = 0; i < _tools.Length; i++)
        {
            _tools[i] = tArray[i];
        }
    }
    IEnumerator IEnumerable.GetEnumerator()
    {
        return (IEnumerator)GetEnumerator();
    }

    public ToolsEnum GetEnumerator()
    {
        return new ToolsEnum(_tools);
    }

}

public class ToolsEnum : IEnumerator
{
    public Tool[] _tools;
    int position = -1;

    public ToolsEnum(Tool[] list)
    {
        _tools = list;
    }
    public bool MoveNext()
    {
        position++;
        return position >= _tools.Length;
    }
    public void Reset()
    {
        position = -1;
    }
    object IEnumerator.Current
    {
        get
        {
            return Current;
        }
    }
    public Tool Current
    {
        get
        {
            try
            {
                return _tools[position];
            }
            catch (IndexOutOfRangeException)
            {
                throw new IndexOutOfRangeException();
            }
        }
    }
}
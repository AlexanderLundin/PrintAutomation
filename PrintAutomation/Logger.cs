using System;
using System.IO;

public partial class TextFileLog
{

    private string _logFilePath = "";

    public TextFileLog()
    {
        _logFilePath = Path.GetTempFileName();
    }

    public TextFileLog(string filePath)
    {
        string dir = Path.GetDirectoryName(filePath);
        string fileName = Path.GetFileNameWithoutExtension(filePath);
        string fileExt = Path.GetExtension(filePath);

        // create directory if it does not exist
        if (!Directory.Exists(dir))
        {
            var di = Directory.CreateDirectory(dir);
        }
        _logFilePath = string.Concat(dir, @"\", fileName, fileExt);
    }


    public bool Write(string message)
    {
        try
        {
            string timeStamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            // pad to 19 spaces to keep log file lined up
            timeStamp = timeStamp.PadRight(19);
            File.AppendAllText(_logFilePath, Environment.NewLine + timeStamp + " =>" + "\t" + message);
            return true;
        }
        catch (Exception e)
        {
            return false;
        }
        finally
        {
        }
    }

    public void Exception(Exception ex)
    {
        Write("-----------------------------------------------------------------------------");
        if (ex != null)
        {
            Write(ex.GetType().FullName);
            Write("Message : ".PadRight(26) + ex.Message);
            Write("StackTrace : ".PadRight(26) + ex.StackTrace);
            Write("InnerException Message : " + ex.InnerException.Message);
        }
        Write("-----------------------------------------------------------------------------");
    }

    public void Delete()
    {
        string dir = Path.GetDirectoryName(_logFilePath);
        if (Directory.Exists(dir))
        {
            if (File.Exists(_logFilePath))
            {
                File.Delete(_logFilePath);
            }
        }
    }


}
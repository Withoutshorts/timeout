<%@ Page Language="C#" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

/// <summary>
/// Summary description for GoogleMap
/// </summary>
public class GoogleMap
{
    public string Status{get;set;}
    public string OriginAddress{get;set;}
    public string DestinationAddress { get; set; }
    public Int64 DurationSeconds{get;set;}
    public string DurationText { get; set; }
    public Int64 DistanceMeters { get; set; }
    public string DistanceText { get; set; }

	public GoogleMap()
	{
		//
		// TODO: Add constructor logic here
		//
	}

    public static GoogleMap Read(string uri)
    {
        GoogleMap gm = new GoogleMap();

        using (System.Xml.XmlReader xmlReader1 = System.Xml.XmlReader.Create(uri))
        {
            while (xmlReader1.Read())
            {                
                if(xmlReader1.NodeType == System.Xml.XmlNodeType.Element)
                {                   
                    if (xmlReader1.Name == "status")
                    {
                        xmlReader1.Read();
                        if(xmlReader1.NodeType == System.Xml.XmlNodeType.Text)
                            gm.Status = xmlReader1.Value;
                    }
                    else if (xmlReader1.Name == "origin_address")
                    {
                        xmlReader1.Read();
                        if (xmlReader1.NodeType == System.Xml.XmlNodeType.Text)
                            gm.OriginAddress = xmlReader1.Value;
                    }
                    else if (xmlReader1.Name == "destination_address")
                    {
                        xmlReader1.Read();
                        if (xmlReader1.NodeType == System.Xml.XmlNodeType.Text)
                            gm.DestinationAddress = xmlReader1.Value;
                    }
                    else if (xmlReader1.Name == "duration")
                    {
                        if (xmlReader1.ReadToDescendant("value"))
                        {
                             xmlReader1.Read();
                             if (xmlReader1.NodeType == System.Xml.XmlNodeType.Text)
                             {
                                 Int64 sec = 0;
                                 Int64.TryParse(xmlReader1.Value, out sec);
                                 gm.DurationSeconds = sec;

                                 xmlReader1.Read(); //read end element
                             }
                        }

                        if (xmlReader1.ReadToNextSibling("text"))
                        {
                             xmlReader1.Read();
                             if (xmlReader1.NodeType == System.Xml.XmlNodeType.Text)
                             {
                                 gm.DurationText = xmlReader1.Value;
                             }
                        }
                    }
                    else if (xmlReader1.Name == "distance")
                    {
                        if (xmlReader1.ReadToDescendant("value"))
                        {
                            xmlReader1.Read();
                            if (xmlReader1.NodeType == System.Xml.XmlNodeType.Text)
                            {
                                Int64 meter = 0;
                                Int64.TryParse(xmlReader1.Value, out meter);
                                gm.DistanceMeters = meter;

                                xmlReader1.Read();  //read end element
                            }
                        }

                        if (xmlReader1.ReadToNextSibling("text"))
                        {
                            xmlReader1.Read();
                            if (xmlReader1.NodeType == System.Xml.XmlNodeType.Text)
                            {
                                gm.DistanceText = xmlReader1.Value;
                            }
                        }
                    }                   
                }                
            }
        }

        return gm;
    }

    public static string ConvertAddress(string address)
    {
        string convertedAddress = string.Empty;

        address = address.Replace("\r", "");

        char[] seperators = {' ', ',','\n'};
        string[] data = address.Split(seperators);
        for (int i = 0; i < data.Length; i++)
        {
            if(data[i].Trim() != string.Empty)
                convertedAddress += data[i] + "+";
        }

        convertedAddress = convertedAddress.Substring(0, convertedAddress.Length-1);

        return convertedAddress;
    }

    public static string[] SeperateAddress(string strAddress)
    {
        string[] addr = new string[2];

        int indexFrom = 0;
        int indexTo = 0;
        strAddress = HttpUtility.UrlDecode(strAddress);
        if (strAddress !=null && strAddress.Contains("Fra"))
        {
            indexFrom = strAddress.IndexOf("Fra:");
            indexTo = strAddress.IndexOf("Til:");

            addr[0] = strAddress.Substring(indexFrom + 4, indexTo - (indexFrom + 4));
            addr[1] = strAddress.Substring(indexTo + 4);  
        }
        else
        {            
            indexFrom = strAddress.IndexOf("From:");
            indexTo = strAddress.IndexOf("To:");
            
            addr[0] = strAddress.Substring(indexFrom + 5, indexTo-(indexFrom + 5));
            addr[1] = strAddress.Substring(indexTo + 3);  
        }             

        return addr;
    }
}
    protected void Page_Load(object sender, EventArgs e)
    {
        try
        {
            if (!string.IsNullOrEmpty(Request["comments"]))
            {
                string[] addr = GoogleMap.SeperateAddress(Request["comments"]);
                addr[0] = GoogleMap.ConvertAddress(addr[0]);
                addr[1] = GoogleMap.ConvertAddress(addr[1]);
                GoogleMap googleData = GoogleMap.Read("http://maps.googleapis.com/maps/api/distancematrix/xml?origins=" + addr[0] + "&destinations=" + addr[1] + "&mode=driving&language=da-DK&sensor=false");

                lblStatus.Text = "<b>Fra</b>: " + googleData.OriginAddress + "<br><b>Til</b>: " + googleData.DestinationAddress + "<br><b>Distance</b>: ";
                lblKM.Text = googleData.DistanceText;
            }
            else
                lblKM.Text = "0 km";
        }
        catch (Exception ex)
        {
            lblStatus.Text = "Server error: " + ex.Message;
        }
    }
    
    
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <script type="text/javascript" src="inc/timereg_2012_4_func.js"></script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
     <h2>Google Map API</h2>
        <asp:Label ID="lblStatus" runat="server" Text="Calculating distance..."></asp:Label>
        <asp:Label ID="lblKM" runat="server" Text=""></asp:Label>
        <br />

        <div style="text-align:center;">
        <br />
        <asp:Button ID="btnClose" runat="server" Text="OK" OnClientClick="SetValueInParent();" />
        </div>
    </div>
    </form>
</body>
</html>

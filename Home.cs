using YamlDotNet.Serialization;

namespace Home;
public class Home
{
    public string World { get; set; } = string.Empty;
    [YamlMember(Alias = "world-name")]
    public string WorldName { get; set; } = string.Empty;
    public double X { get; set; }
    public double Y { get; set; }
    public double Z { get; set; }
    public double Yaw { get; set; }
    public double Pitch { get; set; }
}

public class UserData
{
    [YamlMember(Alias = "last-account-name")]
    public string LastAccountName { get; set; } = string.Empty;
    public Dictionary<string, Home> Homes { get; set; } = new();
}

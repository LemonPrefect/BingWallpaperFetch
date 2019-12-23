using Flurl.Http;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace BingWallpaperFetch{
    class Program{
        public static string[] resolutions = {
            "1920x1080",
            "1366x768",
            "1280x768",
            "1024x768",
            "800x600",
            "800x480",
            "768x1280",
            "720x1280",
            "640x480",
            "480x800",
            "400x240",
            "320x240",
            "240x320"
        };

        public static string[] regionId = new[]{"cn", "tw", "jp", "us"};

        public struct Wallpaper{
            public string title;
            public string timeStart;
            public string timeEnd;
            public string copyright;
            public string hsh;
            public string compareString;
            public string urlBase;
            public string fetchDate;
            public string deleteArg;
        }
        public struct Config{
            public string ver;
            public string wallpaperFileRootPath;
        }

        public static Encoding GB2312 = Encoding.GetEncoding(936);

        public static bool SetImageProperty(Image image, byte[] content, int contentId, short propertyType){
            try{
                PropertyItem propertyInf = image.PropertyItems[0];
                propertyInf.Id = contentId;
                propertyInf.Type = propertyType;
                propertyInf.Value = content;
                propertyInf.Len = propertyInf.Value.Length;
                image.SetPropertyItem(propertyInf);
                return true;
            }
            catch{
                ShowPopUpNotification("Set wallpaper PropertyItem failed", "Warning", ToolTipIcon.Warning);
                return false;
            }
        }
        public static bool ShowPopUpNotification(string content, string title, ToolTipIcon iconType){
            NotifyIcon annouceIcon = new NotifyIcon {
                Icon = SystemIcons.Exclamation,
                BalloonTipTitle = title,
                BalloonTipText = content,
                BalloonTipIcon = iconType,
                Visible = true
            };
            annouceIcon.ShowBalloonTip(1);
            annouceIcon.Visible = false;
            annouceIcon.Dispose();
            if (iconType == ToolTipIcon.Error){
                Thread.CurrentThread.Abort();
            }
            return true;
        }

        public static bool WallpaperWebFetch(int idx, ref Wallpaper[] wallpaperFetched, string[] regionIds){
            try{
                int countNewWallpaper = 0;
                for (int i = 0; i < regionIds.Length; i++, countNewWallpaper++){
                    string response;
                    if (regionIds[i] == "cn"){
                        response = ("https://cn.bing.com/HPImageArchive.aspx?format=js&idx=" + idx + "&n=1&ensearch=0")
                            .WithHeader("X-Forwarded-For", "220.181.38.148").WithCookie("ENSEARCH", "BENVER=1")
                            .GetStringAsync().Result;
                    }
                    else{
                        response =
                            ("https://cn.bing.com/HPImageArchive.aspx?format=js&idx=" + idx + "&n=1&cc=" + regionIds[i])
                            .WithHeader("X-Forwarded-For", "64.233.161.2").WithCookie("ENSEARCH", "BENVER=1")
                            .GetStringAsync().Result;
                    }

                    JObject responseJson = JObject.Parse(response);
                    JArray images = (JArray) responseJson["images"];
                    int flag = 0;
                    //获取新壁纸去重【标题贪心性】
                    for (int j = 0; j < i; j++){
                        string tempCompareString = ((string) images[0]["urlbase"])
                            .Substring(((string) images[0]["urlbase"])
                                .IndexOf("OHR.", StringComparison.Ordinal)
                                , ((string) images[0]["urlbase"])
                                  .IndexOf("_", StringComparison.Ordinal) - ((string) images[0]["urlbase"])
                                  .IndexOf("OHR.", StringComparison.Ordinal));

                        if (tempCompareString == wallpaperFetched[j].compareString){
                            if (wallpaperFetched[j].title == ""){
                                wallpaperFetched[j].title = (string) images[0]["title"];
                            }

                            flag = 1;
                            break;
                        }
                    } 
                    
                    if (flag == 1){
                        continue;
                    }

                    wallpaperFetched[countNewWallpaper] = new Wallpaper {
                        title = (string) images[0]["title"],
                        timeStart = (string) images[0]["startdate"],
                        timeEnd = (string) images[0]["enddate"],
                        urlBase = (string) images[0]["urlbase"],
                        copyright = (string) images[0]["copyright"],
                        hsh = (string) images[0]["hsh"],
                        fetchDate = (string) images[0]["startdate"],
                    };
                    wallpaperFetched[countNewWallpaper].compareString = wallpaperFetched[countNewWallpaper].urlBase
                        .Substring(wallpaperFetched[countNewWallpaper].urlBase
                                .IndexOf("OHR.", StringComparison.Ordinal)
                            , wallpaperFetched[countNewWallpaper].urlBase
                                  .IndexOf("_", StringComparison.Ordinal) - (wallpaperFetched[countNewWallpaper].urlBase)
                              .IndexOf("OHR.", StringComparison.Ordinal));
                    wallpaperFetched[countNewWallpaper].deleteArg = "["
                                                                    + wallpaperFetched[countNewWallpaper].timeStart
                                                                    + "-"
                                                                    + wallpaperFetched[countNewWallpaper].timeEnd
                                                                    + "]"
                                                                    + wallpaperFetched[countNewWallpaper].compareString
                                                                    + "-"
                                                                    + wallpaperFetched[countNewWallpaper].hsh;
                }
            }
            catch{
                ShowPopUpNotification("[WebRequest]Fail to fetch wallpaer and create an object", "Error",
                    ToolTipIcon.Error);
                Thread.CurrentThread.Abort();
                return false;
            }
            return true;
        }

        public static bool ReadComaprePool(string comparePoolPath, ref Wallpaper[] localWallpapers){
            if (File.Exists(comparePoolPath) == true){
                try{
                    StreamReader comparePoolStream = File.OpenText(comparePoolPath);
                    JsonTextReader comparePoolJsonReader = new JsonTextReader(comparePoolStream);
                    JObject comparePool = (JObject) JToken.ReadFrom(comparePoolJsonReader);
                    comparePoolStream.Close();
                    Wallpaper[] wallpapers = new Wallpaper[((JArray) comparePool["items"]).Count];
                    int countValidWallpaperByTime = 0;
                    for (int i = 0; i < ((JArray) comparePool["items"]).Count; i++){
                        Console.WriteLine();
                        TimeSpan itemSpan = Convert.ToDateTime(DateTime.Now).Subtract(
                            DateTime.ParseExact((string) ((JArray) comparePool["items"])[i]["fetchDate"], "yyyyMMdd",
                                System.Globalization.CultureInfo.InvariantCulture));
                        if (itemSpan.Days >= 10){
                            continue;
                        }
                        
                        wallpapers[countValidWallpaperByTime] = new Wallpaper {
                            title = (string) ((JArray) comparePool["items"])[i]["title"],
                            timeStart = (string) ((JArray) comparePool["items"])[i]["timeStart"],
                            timeEnd = (string) ((JArray) comparePool["items"])[i]["timeEnd"],
                            hsh = (string) ((JArray) comparePool["items"])[i]["hsh"],
                            compareString = (string) ((JArray) comparePool["items"])[i]["compareString"],
                            copyright = (string) ((JArray) comparePool["items"])[i]["copyright"],
                            deleteArg = (string) ((JArray) comparePool["items"])[i]["deleteArg"],
                            fetchDate = (string) ((JArray) comparePool["items"])[i]["fetchDate"],
                            urlBase = (string) ((JArray) comparePool["items"])[i]["urlBase"]
                        };
                        countValidWallpaperByTime++;
                    }

                    localWallpapers = wallpapers;
                    return true;
                }
                catch{
                    ShowPopUpNotification("ComaprePool does exits but read failed.", "Error", ToolTipIcon.Error);
                    return false;
                }
            }
            else{
                try{
                    File.Create(comparePoolPath);
                }
                catch{
                    ShowPopUpNotification("[Create File]File path is not correct", "Error", ToolTipIcon.Error);
                    return false;
                }

                return true;
            }
        }

        public static bool DeleteSpecificFile(string fileName, string filePathWithoutSolution){
            try{
                for (int i = 0; i < resolutions.Length; i++){
                    File.Delete(filePathWithoutSolution + resolutions[i] + "\\" + "[" + resolutions[i] + "]" +
                                fileName);
                }

                if (File.Exists(filePathWithoutSolution + "\\1920x1200\\" + "[1920x1200]" + fileName) == true){
                    File.Delete(filePathWithoutSolution + "\\1920x1200\\" + fileName);
                }

                return true;
            }
            catch{
                ShowPopUpNotification("[File Modifiy]File delete failed.", "Error", ToolTipIcon.Error);
                return false;
            }
        }

        public static bool DownloadWallpaper(string url, string fileSavePath, string fileName){
            try{
                string nullpath = url.DownloadFileAsync(fileSavePath, fileName).Result;
            }
            catch{
                ShowPopUpNotification("[Web Request]Download the wallpaper failed.", "Error", ToolTipIcon.Error);
                return false;
            }

            return true;
        }

        public static bool ReadConfig(string configFilePath, ref Config config){
            int flag = 0;
            try{
                StreamReader configStream = File.OpenText("config.json");
                JsonTextReader configJsonReader = new JsonTextReader(configStream);
                JObject configJson = (JObject) JToken.ReadFrom(configJsonReader);
                config = new Config {
                    ver = (string) configJson["ver"],
                    wallpaperFileRootPath = (string) configJson["wallpaperFileRootPath"],
                };
                if (config.ver == ""){
                    config.ver = "1.0.0.0";
                    flag = 1;
                }

                if (config.wallpaperFileRootPath == ""){
                    config.wallpaperFileRootPath = "D:\\BingWallpaper\\";
                    flag = 1;
                }
            }
            catch{
                config = new Config {
                    ver = "1.0.0.0",
                    wallpaperFileRootPath = "D:\\BingWallpaper\\",
                };
                flag = 1;
            }

            if (flag == 1){
                try{
                    dynamic configFile = new JObject();
                    configFile["ver"] = config.ver;
                    configFile["wallpaperFileRootPath"] = config.wallpaperFileRootPath;
                    File.WriteAllText(configFilePath, JsonConvert.SerializeObject(configFile, Formatting.Indented));
                }
                catch{
                    ShowPopUpNotification("[Config Write]Write the config to file error!", "Error", ToolTipIcon.Error);
                    return false;
                }
            }

            return true;
        }

        public static bool TrimWallpaper(ref Wallpaper[] wallpapers){
            try{
                int countValidWallpaper = 0;
                for (int i = 0; i < wallpapers.Length; i++){
                    if (!string.IsNullOrEmpty(wallpapers[i].compareString)){
                        countValidWallpaper++;
                    }
                }

                Wallpaper[] trimmedWallpaper = new Wallpaper[countValidWallpaper];
                int pushWallpaperTrim = 0;
                for (int i = 0; i < wallpapers.Length; i++){
                    if (!string.IsNullOrEmpty(wallpapers[i].compareString)){
                        trimmedWallpaper[pushWallpaperTrim] = wallpapers[i];
                        pushWallpaperTrim++;
                    }
                }

                wallpapers = trimmedWallpaper;
            }
            catch{
                ShowPopUpNotification("[Trim Wallpaper]Trim the new fetched wallpaper failed", "Error",
                    ToolTipIcon.Error);
                return false;
            }

            return true;
        }

        public static bool CheckDirectory(string filePathWithoutSolution){
            try{
                string dirPathStr;
                DirectoryInfo dirPath;
                for (int i = 0; i < resolutions.Length; i++){
                    dirPathStr = filePathWithoutSolution + resolutions[i] + "\\";
                    dirPath = new DirectoryInfo(dirPathStr);
                    if (dirPath.Exists == false){
                        dirPath.Create();
                    }
                }

                dirPathStr = filePathWithoutSolution + "1920x1200" + "\\";
                dirPath = new DirectoryInfo(dirPathStr);
                if (dirPath.Exists == false){
                    dirPath.Create();
                }
            }
            catch{
                ShowPopUpNotification("[Directory Check]Check the file directory error", "Error", ToolTipIcon.Error);
                return false;
            }

            return true;
        }

        public static bool WriteComparePool(Wallpaper[] newWallpapers, Wallpaper[] oldWallpapers,
            string comparePoolPath){
            try{
                Wallpaper[] items = new Wallpaper[newWallpapers.Length + oldWallpapers.Length];
                int itemLength = items.Length;
                for (int i = 0; i < itemLength; i++){
                    if (i < newWallpapers.Length){
                        items[i] = newWallpapers[i];
                    }
                    else{
                        items[i] = oldWallpapers[i - newWallpapers.Length];
                    }
                }

                dynamic newComparePool = new JObject{{"Version", Application.ProductVersion}};
                newComparePool.Add("items", JsonConvert.SerializeObject(items, Formatting.Indented).Replace("\r\n", "")
                    .Replace(" ", ""));
                File.WriteAllText(comparePoolPath, JsonConvert.SerializeObject(newComparePool, Formatting.Indented));
                string comparePoolText = File.ReadAllText(comparePoolPath);
                for (int i = 0; i < items.Length; i++){
                    if (items[i].title != ""){
                        comparePoolText = comparePoolText.Replace(items[i].title.Replace(" ", ""), items[i].title);
                    }
                }

                File.WriteAllText(comparePoolPath, comparePoolText.Replace("\\", "")
                    .Replace("\"[{", "[{")
                    .Replace("}]\"", "}]"));
            }
            catch{
                ShowPopUpNotification("[Write ComaprePool]Construct and write comaprePool error", "Error",
                    ToolTipIcon.Error);
                return false;
            }

            return true;
        }

        static void Main(string[] args){
            int idx = 0;
            Config config = new Config();
            ReadConfig("config.json", ref config);
            if (args.Length == 2){
                if ((args[0] == "/delete" || args[0] == "-delete") && args.Length == 2 &&
                    (args[0] != "-idx" || args[0] != "/idx") == true){
                    DeleteSpecificFile(args[1], config.wallpaperFileRootPath);
                    return;
                }

                if ((args[0] == "/idx" || args[0] == "-idx") && args.Length == 2 &&
                    (args[0] != "-delete" || args[0] != "/delete") == true && args[1].Length < 3){
                    idx = Convert.ToInt32(args[1]);
                }
            }

            Wallpaper[] newWallpapers = new Wallpaper[regionId.Length];
            WallpaperWebFetch(idx, ref newWallpapers, regionId);
            TrimWallpaper(ref newWallpapers);
            Wallpaper[] localWallpapers = new Wallpaper[50];
            ReadComaprePool(config.wallpaperFileRootPath + "compare.pool", ref localWallpapers);
            //
            //Console.WriteLine(newWallpapers.Length);
            for (int i = 0; i < newWallpapers.Length; i++){
                int flag = 0;
                for (int j = 0; j < localWallpapers.Length; j++){
                    if (newWallpapers[i].compareString == localWallpapers[j].compareString){
                        if (newWallpapers[i].title != "" && localWallpapers[j].title == ""){
                            localWallpapers[j].title = newWallpapers[i].title;
                            for (int k = 0; k < resolutions.Length; k++){
                                string filePath = config.wallpaperFileRootPath + resolutions[k] + "\\";
                                string fileName = "[" + resolutions[k] + "]" + localWallpapers[j].deleteArg + ".jpg";
                                Image img = Image.FromFile(filePath + fileName);
                                SetImageProperty(img, GB2312.GetBytes(Convert.ToString(newWallpapers[i].title)), 0x010E,
                                    2);
                                img.Save(filePath + "_temp_" + fileName);
                                img.Dispose();
                                File.Delete(filePath + fileName);
                                File.Move(filePath + "_temp_" + fileName, filePath + fileName);
                            }

                            flag = 1;
                        }

                        newWallpapers[i].compareString = "";
                        if (flag == 1){
                            break;
                        }
                    }
                }
            }

            TrimWallpaper(ref newWallpapers);
            CheckDirectory(config.wallpaperFileRootPath);
            for (int i = 0; i < newWallpapers.Length; i++){
                string fileName = " ";
                string filePath = " ";
                for (int j = 0; j < resolutions.Length; j++){
                    fileName = "[" + resolutions[j] + "]"
                               + "[" + newWallpapers[i].timeStart
                               + "-" + newWallpapers[i].timeEnd + "]"
                               + newWallpapers[i].compareString
                               + "-" + newWallpapers[i].hsh + ".jpg";
                    filePath = config.wallpaperFileRootPath + resolutions[j] + "\\";
                    DownloadWallpaper("https://cn.bing.com" + newWallpapers[i].urlBase
                                                            + "_" + resolutions[j] + ".jpg"
                        , filePath
                        , fileName);
                    Image img = Image.FromFile(filePath + fileName);
                    //SetImageProperty(img, Encoding.UTF8.GetBytes(newWallpapers[i].title), 0x01E, 2);
                    SetImageProperty(img, GB2312.GetBytes(Convert.ToString(newWallpapers[i].title)), 0x010E, 2);
                    img.Save(filePath + "_temp_" + fileName);
                    img.Dispose();
                    File.Delete(filePath + fileName);
                    File.Move(filePath + "_temp_" + fileName, filePath + fileName);
                    img = Image.FromFile(filePath + fileName);
                    SetImageProperty(img, GB2312.GetBytes(Convert.ToString(newWallpapers[i].copyright)), 0x8298, 2);
                    img.Save(filePath + "_temp_" + fileName);
                    img.Dispose();
                    File.Delete(filePath + fileName);
                    File.Move(filePath + "_temp_" + fileName, filePath + fileName);
                }

                try{
                    filePath = config.wallpaperFileRootPath + "1920x1200" + "\\";
                    fileName = "[1920x1200]"
                               + "[" + newWallpapers[i].timeStart
                               + "-" + newWallpapers[i].timeEnd + "]"
                               + newWallpapers[i].compareString
                               + "-" + newWallpapers[i].hsh + ".jpg";
                    string nullpath = ("https://cn.bing.com" + newWallpapers[i].urlBase + "_1920x1200.jpg")
                        .DownloadFileAsync(filePath, fileName).Result;
                    /*
                    Image img = Image.FromFile(filePath + fileName);
                    SetImageProperty(img, GB2312.GetBytes(Convert.ToString(newWallpapers[i].title)), 0x010E, 2);
                    img.Save(filePath + "_temp_" + fileName);
                    img.Dispose();
                    File.Delete(filePath + fileName);
                    File.Move(filePath + "_temp_" + fileName, filePath + fileName);
                    img = Image.FromFile(filePath + fileName);
                    SetImageProperty(img, GB2312.GetBytes(Convert.ToString(newWallpapers[i].copyright)), 0x8298, 2);
                    img.Save(filePath + "_temp_" + fileName);
                    img.Dispose();
                    File.Delete(filePath + fileName);
                    File.Move(filePath + "_temp_" + fileName, filePath + fileName);
                    */
                }
                catch{
                    //当获取最大分辨率壁纸失败时，抛空。
                }
            }

            File.Copy(config.wallpaperFileRootPath + "compare.pool", config.wallpaperFileRootPath + "compare.pool.bak",
                true);
            TrimWallpaper(ref localWallpapers);
            if (WriteComparePool(newWallpapers, localWallpapers, config.wallpaperFileRootPath + "compare.pool") ==
                true){
                File.Delete(config.wallpaperFileRootPath + "compare.pool.bak");
            }
            else{
                File.Delete(config.wallpaperFileRootPath + "compare.pool");
                File.Move(config.wallpaperFileRootPath + "compare.pool.bak",
                    config.wallpaperFileRootPath + "compare.pool");
            }

            ShowPopUpNotification("Done", "BingWallpaperFetch", ToolTipIcon.Info);
        }
    }
}
//最后编译通过时间：201912231719
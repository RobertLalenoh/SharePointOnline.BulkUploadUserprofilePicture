using ConsoleApp.Properties;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Net;
using System.Security;

namespace ConsoleApp
{
    class Program
    {
        static string importCsvLocation = Settings.Default.Import_csv_location;
        static string profilePrefix = Settings.Default.Profile_prefix;
        static int sleepTimeMs = Settings.Default.Sleep_time_millisecond;
        static string mysiteUrl = Settings.Default.Mysite_url;
        static string tenantAdminUrl = Settings.Default.Tenant_admin_url;
        static string tenantAdminUsername = Settings.Default.Tenant_admin_username;
        static string tenantAdminPassword = Settings.Default.Tenant_admin_password;
        static int smallThumbWidth = Settings.Default.Small_thumbwidth;
        static int mediumThumbWidth = Settings.Default.Medium_thumbwidth;
        static int largeThumbWidth = Settings.Default.Large_thumbwidth;
        static string relativePathUserprofilePictureLibrary = Settings.Default.Relative_path_userprofile_picture_library;

        static void Main(string[] args)
        {
            try
            {
                int count = 0;
                using (StreamReader readFile = new StreamReader(importCsvLocation))
                {
                    string line;
                    string[] row;
                    string sPoUserProfileName = "";
                    string sourcePictureUrl;

                    while ((line = readFile.ReadLine()) != null)
                    {
                        //ignore first line
                        if (count > 0)
                        {
                            try
                            {
                                row = line.Split(',');
                                sPoUserProfileName = row[0];
                                sourcePictureUrl = row[1];

                                //get source picture from source picture path
                                using (MemoryStream picturefromExchange = GetImagefromHTTPUrl(sourcePictureUrl))
                                {
                                    if (picturefromExchange != null)//if we got picture, upload to SPO
                                    {
                                        //create SP naming convetion for image file
                                        string newImageNamePrefix = sPoUserProfileName.Replace("@", "_").Replace(".", "_");
                                        //upload source image to SPO (might do some resize work, and multiple image upload depending on config file)
                                        string spoImageUrl = UploadImageToSpo(newImageNamePrefix, picturefromExchange);
                                        if (spoImageUrl.Length > 0)//if upload worked
                                        {
                                            SetSingleValueProfileProperty(profilePrefix + sPoUserProfileName, "PictureURL", spoImageUrl);
                                            System.Console.WriteLine("The userprofile picture from user: " + sPoUserProfileName + " has been uploaded.");
                                        }
                                    }
                                }
                                System.Threading.Thread.Sleep(sleepTimeMs);
                            }
                            catch (Exception ex)
                            {
                                System.Console.WriteLine("Error during upload profile-picture from user: " + sPoUserProfileName + " Exception: " + ex.Message);
                            }

                        }
                        count++;
                    }
                }

            }
            catch (Exception ex)
            {
                System.Console.WriteLine("Error during processing CSV file. Exception: " + ex.Message);
                System.Console.ReadKey();
            }
            finally
            {
                System.Console.ReadKey();
            }
        }
        static SecureString GetSecurePassword(string Password)
        {
            SecureString sPassword = new SecureString();
            foreach (char c in Password.ToCharArray()) sPassword.AppendChar(c);
            return sPassword;
        }
        static Stream ResizeImageSmall(Stream OriginalImage, int NewWidth)
        {

            //when resizing large images i.e. bigger than 200px, we lose quality using the GetThumbnailImage method. There are better ways to do this, but will look to imporve in a future version
            // e.g. http://stackoverflow.com/questions/87753/resizing-an-image-without-losing-any-quality
            try
            {
                OriginalImage.Seek(0, SeekOrigin.Begin);
                Image originalImage = Image.FromStream(OriginalImage, true, true);
                if (originalImage.Width == NewWidth) //if sourceimage is same as destination, no point resizing, as it loses quality
                {
                    OriginalImage.Seek(0, SeekOrigin.Begin);
                    originalImage.Dispose();
                    return OriginalImage; //return same image that was passed in
                }
                else
                {
                    Image resizedImage = originalImage.GetThumbnailImage(NewWidth, (NewWidth * originalImage.Height) / originalImage.Width, null, IntPtr.Zero);
                    MemoryStream memStream = new MemoryStream();
                    resizedImage.Save(memStream, ImageFormat.Jpeg);
                    resizedImage.Dispose();
                    originalImage.Dispose();
                    memStream.Seek(0, SeekOrigin.Begin);
                    return memStream;
                }


            }
            catch (Exception ex)
            {
                System.Console.WriteLine(ex.Message);
                return null;
            }
        }
        static Stream ResizeImageLarge(Stream OriginalImage, int NewWidth)
        {
            OriginalImage.Seek(0, SeekOrigin.Begin);
            Image originalImage = Image.FromStream(OriginalImage, true, true);
            int newHeight = (NewWidth * originalImage.Height) / originalImage.Width;

            Bitmap newImage = new Bitmap(NewWidth, newHeight);

            using (Graphics gr = Graphics.FromImage(newImage))
            {
                gr.SmoothingMode = SmoothingMode.HighQuality;
                gr.InterpolationMode = InterpolationMode.HighQualityBicubic;
                gr.PixelOffsetMode = PixelOffsetMode.HighQuality;
                gr.DrawImage(originalImage, new Rectangle(0, 0, NewWidth, newHeight)); //copy to new bitmap
            }


            MemoryStream memStream = new MemoryStream();
            newImage.Save(memStream, ImageFormat.Jpeg);
            originalImage.Dispose();
            memStream.Seek(0, SeekOrigin.Begin);
            return memStream;


        }
        static string UploadImageToSpo(string PictureName, Stream ProfilePicture)
        {
            try
            {
                string spPhotoPathTempate = string.Concat(relativePathUserprofilePictureLibrary, "/{0}_{1}Thumb.jpg"); //path template to photo lib in My Site Host
                string spImageUrl = string.Empty;

                //create SPO Client context to My Site Host
                ClientContext mySiteclientContext = new ClientContext(mysiteUrl);
                SecureString securePassword = GetSecurePassword(tenantAdminPassword);
                //provide auth crendentials using O365 auth
                mySiteclientContext.Credentials = new SharePointOnlineCredentials(tenantAdminUsername, securePassword);                
                using (Stream smallThumb = ResizeImageSmall(ProfilePicture, smallThumbWidth))
                {
                    if (smallThumb != null)
                    {
                        spImageUrl = string.Format(spPhotoPathTempate, PictureName, "S");
                        Microsoft.SharePoint.Client.File.SaveBinaryDirect(mySiteclientContext, spImageUrl, smallThumb, true);
                    }
                }

                //create medium size
                using (Stream mediumThumb = ResizeImageSmall(ProfilePicture, mediumThumbWidth))
                {
                    if (mediumThumb != null)
                    {
                        spImageUrl = string.Format(spPhotoPathTempate, PictureName, "M");
                        Microsoft.SharePoint.Client.File.SaveBinaryDirect(mySiteclientContext, spImageUrl, mediumThumb, true);

                    }
                }

                //create large size image, shown on SkyDrive Pro main page for user
                using (Stream largeThumb = ResizeImageLarge(ProfilePicture, largeThumbWidth))
                {
                    if (largeThumb != null)
                    {

                        spImageUrl = string.Format(spPhotoPathTempate, PictureName, "L");
                        Microsoft.SharePoint.Client.File.SaveBinaryDirect(mySiteclientContext, spImageUrl, largeThumb, true);

                    }
                }
                //return medium sized URL, as this is the one that should be set in the user profile
                return mysiteUrl + string.Format(spPhotoPathTempate, PictureName, "M");
            }
            catch (Exception ex)
            {
                System.Console.WriteLine(ex.Message);
                return string.Empty;
            }

        }
        static MemoryStream GetImagefromHTTPUrl(string imageUrl)
        {
            try
            {
                System.Net.WebRequest webRequest = System.Net.HttpWebRequest.Create(imageUrl);
                WebResponse webResponse = webRequest.GetResponse();
                Stream imageStream = webResponse.GetResponseStream();
                MemoryStream tmpStream = new MemoryStream();
                imageStream.CopyTo(tmpStream);
                tmpStream.Seek(0, SeekOrigin.Begin);
                return tmpStream;
            }
            catch (WebException ex)
            {
                System.Console.WriteLine("Error during making memory stream from picture. Exception: " + ex.Message);

                return null;
            }
            catch (System.Exception ex)
            {
                System.Console.WriteLine("Error during making memory stream from picture. Exception: " + ex.Message);
                return null;
            }
        }
        private static void SetSingleValueProfileProperty(string UserAccountName, string PropertyName, string PropertyValue)
        {
            using (ClientContext clientContext = new ClientContext(tenantAdminUrl))
            {
                SecureString securePassword = GetSecurePassword(tenantAdminPassword);
                clientContext.Credentials = new SharePointOnlineCredentials(tenantAdminUsername, securePassword);
                PeopleManager peopleManager = new PeopleManager(clientContext);
                peopleManager.SetSingleValueProfileProperty(UserAccountName, PropertyName, PropertyValue);
                clientContext.ExecuteQuery();
            }
        }
    }
}

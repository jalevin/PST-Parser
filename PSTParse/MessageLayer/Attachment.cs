using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using PSTParse.ListsTablesPropertiesLayer;

namespace PSTParse.MessageLayer
{
    
    public enum AttachmentMethod
    {
        NONE = 0x00,
        BY_VALUE = 0x01,
        BY_REFERENCE = 0X02,
        BY_REFERENCE_ONLY = 0X04,
        EMBEDDEED_MESSAGE = 0X05,
        STORAGE = 0X06
    }

    
    public class Attachment
    {
        // FIXME don't actually know what this type should be
        public string AttachmentParameters { get; set; } 
        public AttachmentMethod Method { get; set; }
        public uint Size { get; set; }
        public uint RenderingPosition { get; set; }
        public string Filename { get; set; }
        public string FilenameExtension { get; set; }
        public string AttachmentLongFileName { get; set; }
        public string ContentLocation { get; set; }
        public string Pathname { get; set; }
        public string MimeType { get; set; }
        public uint MimeSequence { get; set; }
        public string ContentID { get; set; }
        public string DisplayName { get; set; }
        public uint LTPRowID { get; set; }
        public uint LTPRowVer { get; set; }
        public bool InvisibleInHTML { get; set; }
        public bool InvisibleInRTF { get; set; }
        public bool RenderedInBody { get; set; }
        
        public byte[] Data { get; set; }
        
        public string GUID { get; set; }

        public string GetPropertyString(byte[] data)
        {
            return data == null ? null : Encoding.Unicode.GetString(data);
        }

        public uint GetPropertyUint(byte[] data)
        {
            return data == null ? 0 : BitConverter.ToUInt32(data, 0);
        }
        public void SaveToDisk(string path)
        {
            if(!Directory.Exists(path))
                Directory.CreateDirectory(path);
            
            GUID = Guid.NewGuid().ToString();
            if (Method == AttachmentMethod.EMBEDDEED_MESSAGE)
               Console.WriteLine("Embedded Message" + path + " - " + Method.ToString());
            
            using (Stream file = File.OpenWrite(Path.Combine(path, GUID)))
            {
                file.Write(Data, 0, Data.Length);
            }

        }
        public Attachment(PropertyContext propertyContext) : this(propertyContext?.Properties.Select(d => d.Value)) { }

        public Attachment(IEnumerable<ExchangeProperty> exchangeProperties)
        {
            exchangeProperties = exchangeProperties ?? Enumerable.Empty<ExchangeProperty>();
            foreach (var property in exchangeProperties)
            {
                switch (property.ID)
                {
                    // fixme don't know how to assign this type
                    case MessageProperty.AttachmentParameters:
                        break;
                    case MessageProperty.AttachmentData:
                        // fixme if this is embedded(0x0102) as opposed to binary data, do something else. 
                        Data = property.Data;
                        break;
                    case MessageProperty.AttachmentSize:
                        Size = GetPropertyUint(property.Data);
                        break;
                    case MessageProperty.AttachmentExtension:
                        FilenameExtension = GetPropertyString(property.Data);
                        break;
                    case MessageProperty.AttachmentFileName:
                        if (property.Data != null)
                            Filename = GetPropertyString(property.Data);
                        //else
                        //    Filename = Guid.NewGuid().ToString();
                        break;
                    case MessageProperty.AttachmentPathname:
                        Pathname = GetPropertyString(property.Data);
                        break;
                    case MessageProperty.AttachmentMimeType:
                        MimeType = GetPropertyString(property.Data);
                        break;
                    case MessageProperty.AttachmentMimeSequence:
                        MimeSequence = GetPropertyUint(property.Data);
                        break;
                    case MessageProperty.AttachmentContentID:
                        ContentID = GetPropertyString(property.Data);
                        break;
                    case MessageProperty.AttachmentContentLocation:
                        ContentLocation = GetPropertyString(property.Data);
                        break;
                    case MessageProperty.DisplayName:
                        if (property.Data != null)
                            DisplayName = GetPropertyString(property.Data);
                        else
                            DisplayName = Guid.NewGuid().ToString();
                        break;
                    case MessageProperty.AttachmentLongFileName:
                        if (property.Data != null)
                            AttachmentLongFileName = GetPropertyString(property.Data);
                        //else
                        //    AttachmentLongFileName = Guid.NewGuid().ToString();
                        break;
                    case MessageProperty.AttachmentMethod:
                        Method = (AttachmentMethod)GetPropertyUint(property.Data);
                        break;
                    case MessageProperty.AttachmentRenderPosition:
                        RenderingPosition = GetPropertyUint(property.Data);
                        break;
                    case MessageProperty.AttachmentFlags:
                        var flags = GetPropertyUint(property.Data);
                        InvisibleInHTML = (flags & 0x1) != 0;
                        InvisibleInRTF = (flags & 0x02) != 0;
                        RenderedInBody = (flags & 0x04) != 0;
                        break;
                    case MessageProperty.AttachmentLTPRowID:
                        LTPRowID = GetPropertyUint(property.Data);
                        break;
                    case MessageProperty.AttachmentLTPRowVer:
                        LTPRowVer = GetPropertyUint(property.Data);
                        break;
                    default:
                        break;
                }
            }
        }
    }
}
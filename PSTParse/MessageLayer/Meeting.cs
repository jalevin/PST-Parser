using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using PSTParse.ListsTablesPropertiesLayer;
using PSTParse.NodeDatabaseLayer;
using PSTParse.Utilities;
using static PSTParse.Utilities.Utilities;

namespace PSTParse.MessageLayer
{
   
    public class Meeting : MapiItem 
    {
        private readonly PSTFile _pst;
        private readonly ulong _nid;
        private readonly Lazy<bool> _isRMSEncryptedLazy;
        private readonly Lazy<bool> _isRMSEncryptedHeadersLazy;
        
        public bool IsRMSEncrypted => _isRMSEncryptedLazy.Value;
        public bool IsRMSEncryptedHeaders => _isRMSEncryptedHeadersLazy.Value;
        
        public Recipients Recipients => Lazy(ref _recipientsLazy, GetRecipients);
        public List<Attachment> Attachments => Lazy(ref _attachmentsLazy, GetAttachments);
        public IEnumerable<Attachment> AttachmentHeaders => Lazy(ref _attachmentHeadersLazy, GetAttachmentHeaders);

        private Dictionary<ulong, NodeDataDTO> SubNodeDataDto => Lazy(ref _subNodeDataDtoLazy, () => BlockBO.GetSubNodeData(_nid, _pst));
        private Dictionary<ulong, NodeDataDTO> SubNodeHeaderDataDto => Lazy(ref _subNodeHeaderDataDtoLazy, () => BlockBO.GetSubNodeData(_nid, _pst, take: 1));
        
        private Recipients GetRecipients()
        {
            const ulong recipientsFlag = 1682;
            var recipients = new Recipients();

            var exists = SubNodeDataDto.TryGetValue(recipientsFlag, out NodeDataDTO subNode);
            if (!exists) return recipients;

            var recipientTable = new TableContext(subNode);
            foreach (var row in recipientTable.RowMatrix.Rows)
            {
                var recipient = new Recipient(row);
                switch (recipient.Type)
                {
                    case Recipient.RecipientType.TO:
                        recipients.To.Add(recipient);
                        break;
                    case Recipient.RecipientType.CC:
                        recipients.CC.Add(recipient);
                        break;
                    case Recipient.RecipientType.BCC:
                        recipients.BCC.Add(recipient);
                        break;
                }
            }
            return recipients;
        }

        private IEnumerable<Attachment> GetAttachmentHeaders()
        {
            if (!HasAttachments) yield break;

            var attachmentSet = new HashSet<string>();
            foreach (var subNode in SubNodeHeaderDataDto)
            {
                if ((NodeValue)subNode.Key != NodeValue.AttachmentTable)
                {
                    throw new Exception("expected node to be an attachment table");
                }

                var attachmentTable = new TableContext(subNode.Value);
                var attachmentRows = attachmentTable.RowMatrix.Rows;

                foreach (var attachmentRow in attachmentRows)
                {
                    var attachment = new Attachment(attachmentRow);
                    yield return attachment;
                }
            }
        }

        private List<Attachment> GetAttachments()
        {
            var attachments = new List<Attachment>();
            if (!HasAttachments) return attachments;

            var attachmentSet = new HashSet<string>();
            foreach (var subNode in SubNodeDataDto)
            {
                var nodeType = NID.GetNodeType(subNode.Key);
                if (nodeType != NID.NodeType.ATTACHMENT_PC) continue;

                var attachmentPC = new PropertyContext(subNode.Value);
                var attachment = new Attachment(attachmentPC);
                if (attachmentSet.Contains(attachment.AttachmentLongFileName))
                {
                    var smallGuid = Guid.NewGuid().ToString().Substring(0, 5);
                    attachment.AttachmentLongFileName = $"{smallGuid}-{attachment.AttachmentLongFileName}";
                }
                attachmentSet.Add(attachment.AttachmentLongFileName);
                attachments.Add(attachment);
            }
            return attachments;
        }

        public Meeting(PSTFile pst, PropertyContext propertyContext) : base(pst, propertyContext)
        {
            _nid = propertyContext.NID;
            _pst = pst; 
            _isRMSEncryptedLazy = new Lazy<bool>(GetIsRMSEncrypted);
            _isRMSEncryptedHeadersLazy = new Lazy<bool>(GetIsRMSEncryptedHeaders);
            //_subNodeDataDtoLazy = new Lazy<Dictionary<ulong, NodeDataDTO>>(() => BlockBO.GetSubNodeData(_nid, _pst));
            
            foreach (var prop in PropertyContext.Properties)
            {
                if (prop.Value.Data == null)
                    continue;
                switch (prop.Key)
                {
                    case MessageProperty.ClientSubmitTime:
                        ClientSubmitTime = DateTime.FromFileTimeUtc(BitConverter.ToInt64(prop.Value.Data, 0));
                        break;
                    case MessageProperty.MessageDeliveryTime:
                        MessageDeliveryTime = DateTime.FromFileTimeUtc(BitConverter.ToInt64(prop.Value.Data, 0));
                        break;
                    case MessageProperty.CreationTime:
                        CreationTime = DateTime.FromFileTimeUtc(BitConverter.ToInt64(prop.Value.Data, 0));
                        break;
                    case MessageProperty.LastModificationTime:
                        LastModificationTime = DateTime.FromFileTimeUtc(BitConverter.ToInt64(prop.Value.Data, 0));
                        break;
                    case MessageProperty.MessageSize:
                        MessageSize = BitConverter.ToUInt32(prop.Value.Data, 0);
                        break;
                    case MessageProperty.MessageFlags:
                        MessageFlags = BitConverter.ToUInt32(prop.Value.Data, 0);
                        Read = (MessageFlags & 0x1) != 0;
                        Unsent = (MessageFlags & 0x8) != 0;
                        Unmodified = (MessageFlags & 0x2) != 0;
                        HasAttachments = (MessageFlags & 0x10) != 0;
                        FromMe = (MessageFlags & 0x20) != 0;
                        IsFAI = (MessageFlags & 0x40) != 0;
                        NotifyReadRequested = (MessageFlags & 0x100) != 0;
                        NotifyUnreadRequested = (MessageFlags & 0x200) != 0;
                        EverRead = (MessageFlags & 0x400) != 0;
                        break;
                    case MessageProperty.ConversationTopic:
                        ConversationTopic = Encoding.Unicode.GetString(prop.Value.Data);
                        break;
                    case MessageProperty.Subject:
                        Subject = Encoding.Unicode.GetString(prop.Value.Data);
                        break;
                    case MessageProperty.SenderName:
                        SenderName = Encoding.Unicode.GetString(prop.Value.Data);
                        break;
                    case MessageProperty.SenderAddress:
                        SenderAddress = Encoding.Unicode.GetString(prop.Value.Data);
                        break;
                    case MessageProperty.SenderAddressType:
                        SenderAddressType = Encoding.Unicode.GetString(prop.Value.Data);
                        break;
                    case MessageProperty.Headers:
                        Headers = Encoding.Unicode.GetString(prop.Value.Data);
                        break;
                    case MessageProperty.Importance:
                        Importance = (Importance)BitConverter.ToInt16(prop.Value.Data, 0);
                        break;
                    case MessageProperty.Sensitivity:
                        Sensitivity = (Sensitivity)BitConverter.ToInt16(prop.Value.Data, 0);
                        break;
                    case MessageProperty.Priority:
                        Priority = (Priority)BitConverter.ToInt16(prop.Value.Data, 0);
                        break;
                }
            }
        }
    }

}
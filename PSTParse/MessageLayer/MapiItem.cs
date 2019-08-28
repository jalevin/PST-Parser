using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using PSTParse;
using PSTParse.ListsTablesPropertiesLayer;
using PSTParse.MessageLayer;
using PSTParse.NodeDatabaseLayer;
using PSTParse.Utilities;
using static PSTParse.Utilities.Utilities; 


public class MapiItem : IPMItem
    {
        // must have in every child
        private readonly PSTFile _pst;
        private readonly ulong _nid;
        private readonly Lazy<bool> _isRMSEncryptedLazy;
        private readonly Lazy<bool> _isRMSEncryptedHeadersLazy;
        
        protected Dictionary<ulong, NodeDataDTO> _subNodeDataDtoLazy;
        protected Dictionary<ulong, NodeDataDTO> _subNodeHeaderDataDtoLazy;
        protected Recipients _recipientsLazy;
        protected List<Attachment> _attachmentsLazy;
        protected IEnumerable<Attachment> _attachmentHeadersLazy;
        
        protected Dictionary<ulong, NodeDataDTO> SubNodeDataDto => Lazy(ref _subNodeDataDtoLazy, () => BlockBO.GetSubNodeData(_nid, _pst));
        protected Dictionary<ulong, NodeDataDTO> SubNodeHeaderDataDto => Lazy(ref _subNodeHeaderDataDtoLazy, () => BlockBO.GetSubNodeData(_nid, _pst, take: 1));

        //public bool IsRMSEncrypted => _isRMSEncryptedLazy.Value;
        //public bool IsRMSEncryptedHeaders => _isRMSEncryptedHeadersLazy.Value;
        public Recipients Recipients => Lazy(ref _recipientsLazy, GetRecipients);
        public List<Attachment> Attachments => Lazy(ref _attachmentsLazy, GetAttachments);
        public IEnumerable<Attachment> AttachmentHeaders => Lazy(ref _attachmentHeadersLazy, GetAttachmentHeaders);
       
        // headers
        public DateTime ClientSubmitTime { get; set; } 
        public DateTime ClientDeliveryTime { get; set;}
        public DateTime MessageDeliveryTime { get; set; }
        public DateTime CreationTime { get; set; }   
        public DateTime LastModificationTime { get; set; }
        public uint MessageSize { get; set; }
        public uint MessageFlags { get; set; }
        public string ConversationTopic { get; set; }
        public string Subject { get; set; }
        public string SenderName { get; set; }
        public string SenderAddress { get; set; }
        public string SenderAddressType { get; set; }
        public string SentRepresentingName { get; set; }
        public Importance Importance { get; set; }
        public Priority Priority { get; set; }
        public Sensitivity Sensitivity { get; set; }
        public bool Read { get; set; }
        public bool Unsent { get; set; }
        public bool Unmodified { get; set; }
        public bool HasAttachments { get; set; }
        public bool FromMe { get; set; }
        public bool IsFAI { get; set; }
        public bool NotifyReadRequested { get; set; }
        public bool NotifyUnreadRequested { get; set; }
        public bool EverRead { get; set; }
        public string InternetHeaders { get; set; }
        public string BodyPlainText { get; set; }
        public string SubjectPrefix { get; set; }
        public string BodyHtml { get; set; }
        public uint InternetArticleNumber { get; set; }
        public string BodyCompressedRTFString { get; set; }
        public string InternetMessageID { get; set; }
        public string UrlCompositeName { get; set; }
        public bool AttributeHidden { get; set; }
        public bool ReadOnly { get; set; }
        public uint CodePage { get; set; }
        public uint NonUnicodeCodePage { get; set; }
        
        
        public Recipients GetRecipients()
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
        
        protected IEnumerable<Attachment> GetAttachmentHeaders()
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
        protected List<Attachment> GetAttachments()
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
        protected bool GetIsRMSEncrypted()
        {
            if (!HasAttachments) return false;

            foreach (var attachment in Attachments)
            {
                if (attachment.AttachmentLongFileName?.ToLowerInvariant().EndsWith(".rpmsg") ?? false)
                {
                    if (Attachments.Count > 1) throw new NotSupportedException("too many attachments for rms");
                    return true;
                }
                if (attachment.Filename?.ToLowerInvariant().EndsWith(".rpmsg") ?? false)
                {
                    if (Attachments.Count > 1) throw new NotSupportedException("too many attachments for rms");
                    return true;
                }
                if (attachment.DisplayName?.ToLowerInvariant().EndsWith(".rpmsg") ?? false)
                {
                    if (Attachments.Count > 1) throw new NotSupportedException("too many attachments for rms");
                    return true;
                }
            }
            return false;
        }
        protected bool GetIsRMSEncryptedHeaders()
        {
            if (!HasAttachments) return false;

            foreach (var attachment in AttachmentHeaders)
            {
                Debug.Assert(!(attachment.AttachmentLongFileName?.ToLowerInvariant().EndsWith(".rpm") ?? false));
                Debug.Assert(!(attachment.DisplayName?.ToLowerInvariant().EndsWith(".rpm") ?? false));

                if (attachment.Filename?.ToLowerInvariant().EndsWith(".rpm") ?? false)
                {
                    Debug.Assert(Attachments.Count == 1, "too many attachments for rms");
                    return true;
                }

                return false;
            }
            return false;
        }

        public MapiItem(PSTFile pst, PropertyContext propertyContext) : base(pst, propertyContext)
        {
        }
    }

public enum Priority
{
    NonUrgent = -0x01,
    Normal = 0x00,
    Urgent = 0x01
}
public enum Importance
{
    LOW = 0x00,
    NORMAL = 0x01,
    HIGH = 0x02
}

public enum Sensitivity
{
    Normal = 0x00,
    Personal = 0x01,
    Private = 0x02,
    Confidential = 0x03
}
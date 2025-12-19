

#include "ReportMemoryLeaks.h"

#include "stdafx.h"
#include "property.h"

static CFBProperty propArr[] = {
    { 0x0001 , "PidTagTemplateData" },
    { 0x0002 , "PidTagAlternateRecipientAllowed" },
    { 0x0003 , "PidTagNativeBody" },   // see 0x1016
    { 0x0004 , "PidTagScriptData" },
    { 0x0004 , "PidTagAutoForwardComment" },
    { 0x0005 , "PidTagAutoForwarded" },
    { 0x000F , "PidTagDeferredDeliveryTime" },
    { 0x0010 , "PidTagDeliverTime" },
    { 0x0015 , "PidTagExpiryTime" },
    { 0x0017 , "PidTagImportance" },
    { 0x001A , "PidTagMessageClass" },
    { 0x0023 , "PidTagOriginatorDeliveryReportRequested" },
    { 0x0025 , "PidTagParentKey" },
    { 0x0026 , "PidTagPriority" },
    { 0x0029 , "PidTagReadReceiptRequested" },
    { 0x002A , "PidTagReceiptTime" },
    { 0x002B , "PidTagRecipientReassignmentProhibited" },
    { 0x002E , "PidTagOriginalSensitivity" },
    { 0x0030 , "PidTagReplyTime" },
    { 0x0031 , "PidTagReportTag" },
    { 0x0032 , "PidTagReportTime" },
    { 0x0036 , "PidTagSensitivity" },
    { 0x0037 , "PidTagSubject" },
    { 0x0039 , "PidTagClientSubmitTime" },
    { 0x003A , "PidTagReportName" },
    { 0x003B , "PidTagSentRepresentingSearchKey" },
    { 0x003D , "PidTagSubjectPrefix" },
    { 0x003F , "PidTagReceivedByEntryId" },
    { 0x0040 , "PidTagReceivedByName" },
    { 0x0041 , "PidTagSentRepresentingEntryId" },
    { 0x0042 , "PidTagSentRepresentingName" },
    { 0x0043 , "PidTagReceivedRepresentingEntryId" },
    { 0x0044 , "PidTagReceivedRepresentingName" },
    { 0x0045 , "PidTagReportEntryId" },
    { 0x0046 , "PidTagReadReceiptEntryId" },
    { 0x0047 , "PidTagMessageSubmissionId" },
    { 0x0049 , "PidTagOriginalSubject" },
    { 0x004B , "PidTagOriginalMessageClass" },
    { 0x004C , "PidTagOriginalAuthorEntryId" },
    { 0x004D , "PidTagOriginalAuthorName" },
    { 0x004E , "PidTagOriginalSubmitTime" },
    { 0x004F , "PidTagReplyRecipientEntries" },
    { 0x0050 , "PidTagReplyRecipientNames" },
    { 0x0051 , "PidTagReceivedBySearchKey" },
    { 0x0052 , "PidTagReceivedRepresentingSearchKey" },
    { 0x0053 , "PidTagReadReceiptSearchKey" },
    { 0x0054 , "PidTagReportSearchKey" },
    { 0x0055 , "PidTagOriginalDeliveryTime" },
    { 0x0057 , "PidTagMessageToMe" },
    { 0x0058 , "PidTagMessageCcMe" },
    { 0x0059 , "PidTagMessageRecipientMe" },
    { 0x005A , "PidTagOriginalSenderName" },
    { 0x005B , "PidTagOriginalSenderEntryId" },
    { 0x005C , "PidTagOriginalSenderSearchKey" },
    { 0x005D , "PidTagOriginalSentRepresentingName" },
    { 0x005E , "PidTagOriginalSentRepresentingEntryId" },
    { 0x005F , "PidTagOriginalSentRepresentingSearchKey" },
    { 0x0060 , "PidTagStartDate" },
    { 0x0061 , "PidTagEndDate" },
    { 0x0062 , "PidTagOwnerAppointmentId" },
    { 0x0063 , "PidTagResponseRequested" },
    { 0x0064 , "PidTagSentRepresentingAddressType" },
    { 0x0065 , "PidTagSentRepresentingEmailAddress" },
    { 0x0066 , "PidTagOriginalSenderAddressType" },
    { 0x0067 , "PidTagOriginalSenderEmailAddress" },
    { 0x0068 , "PidTagOriginalSentRepresentingAddressType" },
    { 0x0069 , "PidTagOriginalSentRepresentingEmailAddress" },
    { 0x0070 , "PidTagConversationTopic" },
    { 0x0071 , "PidTagConversationIndex" },
    { 0x0072 , "PidTagOriginalDisplayBcc" },
    { 0x0073 , "PidTagOriginalDisplayCc" },
    { 0x0074 , "PidTagOriginalDisplayTo" },
    { 0x0075 , "PidTagReceivedByAddressType" },
    { 0x0076 , "PidTagReceivedByEmailAddress" },
    { 0x0077 , "PidTagReceivedRepresentingAddressType" },
    { 0x0078 , "PidTagReceivedRepresentingEmailAddress" },
    { 0x007D , "PidTagTransportMessageHeaders" },
    { 0x007F , "PidTagTnefCorrelationKey" },
    { 0x0080 , "PidTagReportDisposition" },
    { 0x0081 , "PidTagReportDispositionMode" },
    { 0x0807 , "PidTagAddressBookRoomCapacity" },
    { 0x0809 , "PidTagAddressBookRoomDescription" },
    { 0x0C04 , "PidTagNonDeliveryReportReasonCode" },
    { 0x0C05 , "PidTagNonDeliveryReportDiagCode" },
    { 0x0C06 , "PidTagNonReceiptNotificationRequested" },
    { 0x0C08 , "PidTagOriginatorNonDeliveryReportRequested" },
    { 0x0C15 , "PidTagRecipientType" },
    { 0x0C17 , "PidTagReplyRequested" },
    { 0x0C19 , "PidTagSenderEntryId" },
    { 0x0C1A , "PidTagSenderName" },
    { 0x0C1B , "PidTagSupplementaryInfo" },
    { 0x0C1D , "PidTagSenderSearchKey" },
    { 0x0C1E , "PidTagSenderAddressType" },
    { 0x0C1F , "PidTagSenderEmailAddress" },
    { 0x0C20 , "PidTagNonDeliveryReportStatusCode" },
    { 0x0C21 , "PidTagRemoteMessageTransferAgent" },
    { 0x0E01 , "PidTagDeleteAfterSubmit" },
    { 0x0E02 , "PidTagDisplayBcc" },
    { 0x0E03 , "PidTagDisplayCc" },
    { 0x0E04 , "PidTagDisplayTo" },
    //{ 0x0E05 , "CodePage" },  // 
    { 0x0E06 , "PidTagMessageDeliveryTime" },
    { 0x0E07 , "PidTagMessageFlags" },
    { 0x0E08 , "PidTagMessageSize" },
    { 0x0E08 , "PidTagMessageSizeExtended" },
    { 0x0E09 , "PidTagParentEntryId" },
    { 0x0E0F , "PidTagResponsibility" },
    { 0x0E12 , "PidTagMessageRecipients" },
    { 0x0E13 , "PidTagMessageAttachments" },
    { 0x0E17 , "PidTagMessageStatus" },
    { 0x0E1B , "PidTagHasAttachments" },
    { 0x0E1D , "PidTagNormalizedSubject" },
    { 0x0E1F , "PidTagRtfInSync" },
    { 0x0E20 , "PidTagAttachSize" },
    { 0x0E21 , "PidTagAttachNumber" },
    { 0x0E28 , "PidTagPrimarySendAccount" },
    { 0x0E29 , "PidTagNextSendAcct" },
    { 0x0E2B , "PidTagToDoItemFlags" },
    { 0x0E2C , "PidTagSwappedToDoStore" },
    { 0x0E2D , "PidTagSwappedToDoData" },
    { 0x0E69 , "PidTagRead" },
    { 0x0E6A , "PidTagSecurityDescriptorAsXml" },
    { 0x0E79 , "PidTagTrustSender" },
    { 0x0E84 , "PidTagExchangeNTSecurityDescriptor" },
    { 0x0E99 , "PidTagExtendedRuleMessageActions" },
    { 0x0E9A , "PidTagExtendedRuleMessageCondition" },
    { 0x0E9B , "PidTagExtendedRuleSizeLimit" },
    { 0x0FF4 , "PidTagAccess" },
    { 0x0FF5 , "PidTagRowType" },
    { 0x0FF6 , "PidTagInstanceKey" },
    { 0x0FF7 , "PidTagAccessLevel" },
    { 0x0FF8 , "PidTagMappingSignature" },
    { 0x0FF9 , "PidTagRecordKey" },
    { 0x0FFB , "PidTagStoreEntryId" },
    { 0x0FFE , "PidTagObjectType" },
    { 0x0FFF , "PidTagEntryId" },
    { 0x1000 , "PidTagBody" },
    { 0x1001 , "PidTagReportText" },
    //{ 0x1002 , "Unknown" },
    { 0x1003 , "PidTagOfflineAddressBookTruncatedProperties" },
    //{ 0x1004 , "Unknown" },
    { 0x1007 , "PidTagRtfSyncBodyCount" },
    { 0x1008 , "PidTagRtfSyncBodyTag" },
    { 0x1009 , "PidTagRtfCompressed" },
    //{ 0x100C , "Unknown" },
    //{ 0x100F , "Unknown" },
    { 0x1010 , "PidTagRtfSyncPrefixCount" },
    { 0x1011 , "PidTagRtfSyncTrailingCount" },
    //{ 0x1012 , "Unknown" },  // CodePage
    //{ 0x1013 , "PidTagHtml" },  // Data type: PtypBinary, 0x0102
    { 0x1013 , "PidTagBodyHtml" }, // Data type: PtypString, 0x001F
    { 0x1014 , "PidTagBodyContentLocation" },
    { 0x1015 , "PidTagBodyContentId" },
    { 0x1016 , "PidTagNativeBody" },
    //{ 0x1017 , "Unknown" },
    //{ 0x1018 , "Unknown" },  // "User Defined"
    //{ 0x1019 , "Unknown" },  // CodePage
    //{ 0x101a , "Unknown" },  // "Last Modified Time"
    //{ 0x101b , "Unknown" },  // CodePage
    //{ 0x101e , "Unknown" },
    { 0x1035 , "PidTagInternetMessageId" },
    { 0x1039 , "PidTagInternetReferences" },
    { 0x1042 , "PidTagInReplyToId" },
    { 0x1043 , "PidTagListHelp" },
    { 0x1044 , "PidTagListSubscribe" },
    { 0x1045 , "PidTagListUnsubscribe" },
    { 0x1046 , "PidTagOriginalMessageId" },
    { 0x1080 , "PidTagIconIndex" },
    { 0x1081 , "PidTagLastVerbExecuted" },
    { 0x1082 , "PidTagLastVerbExecutionTime" },
    { 0x1090 , "PidTagFlagStatus" },
    { 0x1091 , "PidTagFlagCompleteTime" },
    { 0x1095 , "PidTagFollowupIcon" },
    { 0x1096 , "PidTagBlockStatus" },
    { 0x10C3 , "PidTagICalendarStartTime" },
    { 0x10C4 , "PidTagICalendarEndTime" },
    { 0x10C5 , "PidTagCdoRecurrenceid" },
    { 0x10CA , "PidTagICalendarReminderNextTime" },
    { 0x10F4 , "PidTagAttributeHidden" },
    { 0x10F6 , "PidTagAttributeReadOnly" },
    { 0x3000 , "PidTagRowid" },
    { 0x3001 , "PidTagDisplayName" },
    { 0x3002 , "PidTagAddressType" },
    { 0x3003 , "PidTagEmailAddress" },
    { 0x3004 , "PidTagComment" },
    { 0x3005 , "PidTagDepth" },
    { 0x3007 , "PidTagCreationTime" },
    { 0x3008 , "PidTagLastModificationTime" },
    { 0x300B , "PidTagSearchKey" },
    { 0x300F , "OrigEntryId" },
    { 0x3010 , "PidTagTargetEntryId" },
    { 0x3013 , "PidTagConversationId" },
    { 0x3016 , "PidTagConversationIndexTracking" },
    { 0x3018 , "PidTagArchiveTag" },
    { 0x3019 , "PidTagPolicyTag" },
    { 0x301A , "PidTagRetentionPeriod" },
    { 0x301B , "PidTagStartDateEtc" },
    { 0x301C , "PidTagRetentionDate" },
    { 0x301D , "PidTagRetentionFlags" },
    { 0x301E , "PidTagArchivePeriod" },
    { 0x301F , "PidTagArchiveDate" },
    { 0x340D , "PidTagStoreSupportMask" },
    { 0x340E , "PidTagStoreState" },
    { 0x3600 , "PidTagContainerFlags" },
    { 0x3601 , "PidTagFolderType" },
    { 0x3602 , "PidTagContentCount" },
    { 0x3603 , "PidTagContentUnreadCount" },
    { 0x3609 , "PidTagSelectable" },
    { 0x360A , "PidTagSubfolders" },
    { 0x360C , "PidTagAnr" },
    { 0x360E , "PidTagContainerHierarchy" },
    { 0x360F , "PidTagContainerContents" },
    { 0x3610 , "PidTagFolderAssociatedContents" },
    { 0x3613 , "PidTagContainerClass" },
    { 0x36D0 , "PidTagIpmAppointmentEntryId" },
    { 0x36D1 , "PidTagIpmContactEntryId" },
    { 0x36D2 , "PidTagIpmJournalEntryId" },
    { 0x36D3 , "PidTagIpmNoteEntryId" },
    { 0x36D4 , "PidTagIpmTaskEntryId" },
    { 0x36D5 , "PidTagRemindersOnlineEntryId" },
    { 0x36D7 , "PidTagIpmDraftsEntryId" },
    { 0x36D8 , "PidTagAdditionalRenEntryIds" },
    { 0x36D9 , "PidTagAdditionalRenEntryIdsEx" },
    { 0x36DA , "PidTagExtendedFolderFlags" },
    { 0x36E2 , "PidTagOrdinalMost" },
    { 0x36E4 , "PidTagFreeBusyEntryIds" },
    { 0x36E5 , "PidTagDefaultPostMessageClass" },
    { 0x3700 , "PidTagClientActivelyEditingUntil" },
    { 0x3701 , "PidTagAttachDataObject" },
    { 0x3701 , "PidTagAttachDataBinary" },
    { 0x3702 , "PidTagAttachEncoding" },
    { 0x3703 , "PidTagAttachExtension" },
    { 0x3704 , "PidTagAttachFilename" },
    { 0x3705 , "PidTagAttachMethod" },
    { 0x3707 , "PidTagAttachLongFilename" },
    { 0x3708 , "PidTagAttachPathname" },
    { 0x3709 , "PidTagAttachRendering" },
    { 0x370A , "PidTagAttachTag" },
    { 0x370B , "PidTagRenderingPosition" },
    { 0x370C , "PidTagAttachTransportName" },
    { 0x370D , "PidTagAttachLongPathname" },
    { 0x370E , "PidTagAttachMimeTag" },
    { 0x370F , "PidTagAttachAdditionalInformation" },
    { 0x3711 , "PidTagAttachContentBase" },
    { 0x3712 , "PidTagAttachContentId" },
    { 0x3713 , "PidTagAttachContentLocation" },
    { 0x3714 , "PidTagAttachFlags" },
    { 0x3719 , "PidTagAttachPayloadProviderGuidString" },
    { 0x371A , "PidTagAttachPayloadClass" },
    { 0x371B , "PidTagTextAttachmentCharset" },
    //{ 0x371D , "Unknown" },  //  "Last Author"
    { 0x3900 , "PidTagDisplayType" },
    { 0x3902 , "PidTagTemplateid" },
    { 0x3905 , "PidTagDisplayTypeEx" },
    { 0x39FE , "PidTagSmtpAddress" },
    { 0x39FF , "PidTagAddressBookDisplayNamePrintable" },
    { 0x3A00 , "PidTagAccount" },
    { 0x3A02 , "PidTagCallbackTelephoneNumber" },
    { 0x3A05 , "PidTagGeneration" },
    { 0x3A06 , "PidTagGivenName" },
    { 0x3A07 , "PidTagGovernmentIdNumber" },
    { 0x3A08 , "PidTagBusinessTelephoneNumber" },
    { 0x3A09 , "PidTagHomeTelephoneNumber" },
    { 0x3A0A , "PidTagInitials" },
    { 0x3A0B , "PidTagKeyword" },
    { 0x3A0C , "PidTagLanguage" },
    { 0x3A0D , "PidTagLocation" },
    { 0x3A0F , "PidTagMessageHandlingSystemCommonName" },
    { 0x3A10 , "PidTagOrganizationalIdNumber" },
    { 0x3A11 , "PidTagSurname" },
    { 0x3A12 , "PidTagOriginalEntryId" },
    { 0x3A15 , "PidTagPostalAddress" },
    { 0x3A16 , "PidTagCompanyName" },
    { 0x3A17 , "PidTagTitle" },
    { 0x3A18 , "PidTagDepartmentName" },
    { 0x3A19 , "PidTagOfficeLocation" },
    { 0x3A1A , "PidTagPrimaryTelephoneNumber" },
    { 0x3A1B , "PidTagBusiness2TelephoneNumber" },
    { 0x3A1B , "PidTagBusiness2TelephoneNumbers" },
    { 0x3A1C , "PidTagMobileTelephoneNumber" },
    { 0x3A1D , "PidTagRadioTelephoneNumber" },
    { 0x3A1E , "PidTagCarTelephoneNumber" },
    { 0x3A1F , "PidTagOtherTelephoneNumber" },
    { 0x3A20 , "PidTagTransmittableDisplayName" },
    { 0x3A21 , "PidTagPagerTelephoneNumber" },
    { 0x3A22 , "PidTagUserCertificate" },
    { 0x3A23 , "PidTagPrimaryFaxNumber" },
    { 0x3A24 , "PidTagBusinessFaxNumber" },
    { 0x3A25 , "PidTagHomeFaxNumber" },
    { 0x3A26 , "PidTagCountry" },
    { 0x3A27 , "PidTagLocality" },
    { 0x3A28 , "PidTagStateOrProvince" },
    { 0x3A29 , "PidTagStreetAddress" },
    { 0x3A2A , "PidTagPostalCode" },
    { 0x3A2B , "PidTagPostOfficeBox" },
    { 0x3A2C , "PidTagTelexNumber" },
    { 0x3A2D , "PidTagIsdnNumber" },
    { 0x3A2E , "PidTagAssistantTelephoneNumber" },
    { 0x3A2F , "PidTagHome2TelephoneNumber" },
    { 0x3A2F , "PidTagHome2TelephoneNumbers" },
    { 0x3A30 , "PidTagAssistant" },
    { 0x3A40 , "PidTagSendRichInfo" },
    { 0x3A41 , "PidTagWeddingAnniversary" },
    { 0x3A42 , "PidTagBirthday" },
    { 0x3A43 , "PidTagHobbies" },
    { 0x3A44 , "PidTagMiddleName" },
    { 0x3A45 , "PidTagDisplayNamePrefix" },
    { 0x3A46 , "PidTagProfession" },
    { 0x3A47 , "PidTagReferredByName" },
    { 0x3A48 , "PidTagSpouseName" },
    { 0x3A49 , "PidTagComputerNetworkName" },
    { 0x3A4A , "PidTagCustomerId" },
    { 0x3A4B , "PidTagTelecommunicationsDeviceForDeafTelephoneNumber" },
    { 0x3A4C , "PidTagFtpSite" },
    { 0x3A4D , "PidTagGender" },
    { 0x3A4E , "PidTagManagerName" },
    { 0x3A4F , "PidTagNickname" },
    { 0x3A50 , "PidTagPersonalHomePage" },
    { 0x3A51 , "PidTagBusinessHomePage" },
    { 0x3A57 , "PidTagCompanyMainTelephoneNumber" },
    { 0x3A58 , "PidTagChildrensNames" },
    { 0x3A59 , "PidTagHomeAddressCity" },
    { 0x3A5A , "PidTagHomeAddressCountry" },
    { 0x3A5B , "PidTagHomeAddressPostalCode" },
    { 0x3A5C , "PidTagHomeAddressStateOrProvince" },
    { 0x3A5D , "PidTagHomeAddressStreet" },
    { 0x3A5E , "PidTagHomeAddressPostOfficeBox" },
    { 0x3A5F , "PidTagOtherAddressCity" },
    { 0x3A60 , "PidTagOtherAddressCountry" },
    { 0x3A61 , "PidTagOtherAddressPostalCode" },
    { 0x3A62 , "PidTagOtherAddressStateOrProvince" },
    { 0x3A63 , "PidTagOtherAddressStreet" },
    { 0x3A64 , "PidTagOtherAddressPostOfficeBox" },
    { 0x3A70 , "PidTagUserX509Certificate" },
    { 0x3A71 , "PidTagSendInternetEncoding" },
    { 0x3F08 , "PidTagInitialDetailsPane" },
    { 0x3FDE , "PidTagInternetCodepage" },
    { 0x3FDF , "PidTagAutoResponseSuppress" },
    { 0x3FE0 , "PidTagAccessControlListData" },
    { 0x3FE3 , "PidTagDelegatedByRule" },
    { 0x3FE7 , "PidTagResolveMethod" },
    { 0x3FEA , "PidTagHasDeferredActionMessages" },
    { 0x3FEB , "PidTagDeferredSendNumber" },
    { 0x3FEC , "PidTagDeferredSendUnits" },
    { 0x3FED , "PidTagExpiryNumber" },
    { 0x3FEE , "PidTagExpiryUnits" },
    { 0x3FEF , "PidTagDeferredSendTime" },
    { 0x3FF0 , "PidTagConflictEntryId" },
    { 0x3FF1 , "PidTagMessageLocaleId" },
    { 0x3FF8 , "PidTagCreatorName" },
    { 0x3FF9 , "PidTagCreatorEntryId" },
    { 0x3FFA , "PidTagLastModifierName" },
    { 0x3FFB , "PidTagLastModifierEntryId" },
    { 0x3FFD , "PidTagMessageCodepage" },
    { 0x4008 , "MetaTagDnPrefix" },
    { 0x400F , "MetaTagEcWarning" },
    { 0x4011 , "MetaTagNewFXFolder" },
    { 0x4016 , "MetaTagFXDelProp" },
    { 0x4017 , "MetaTagIdsetGiven" },
    { 0x401A , "PidTagSentRepresentingFlags" },
    { 0x4021 , "MetaTagIdsetNoLongerInScope" },
    { 0x4029 , "PidTagReadReceiptAddressType" },
    { 0x402A , "PidTagReadReceiptEmailAddress" },
    { 0x402B , "PidTagReadReceiptName" },
    { 0x402D , "MetaTagIdsetRead" },
    { 0x402E , "MetaTagIdsetUnread" },
    { 0x4076 , "PidTagContentFilterSpamConfidenceLevel" },
    { 0x4079 , "PidTagSenderIdStatus" },
    { 0x407A , "MetaTagIncrementalSyncMessagePartial" },
    { 0x407C , "MetaTagIncrSyncGroupId" },
    { 0x4083 , "PidTagPurportedSenderDomain" },
    { 0x5902 , "PidTagInternetMailOverrideFormat" },
    { 0x5909 , "PidTagMessageEditorFormat" },
    { 0x5D01 , "PidTagSenderSmtpAddress" },
    { 0x5D02 , "PidTagSentRepresentingSmtpAddress" },
    { 0x5D05 , "PidTagReadReceiptSmtpAddress" },
    { 0x5D07 , "PidTagReceivedBySmtpAddress" },
    { 0x5D08 , "PidTagReceivedRepresentingSmtpAddress" },
    { 0x5FDF , "PidTagRecipientOrder" },
    { 0x5FE1 , "PidTagRecipientProposed" },
    { 0x5FE3 , "PidTagRecipientProposedStartTime" },
    { 0x5FE4 , "PidTagRecipientProposedEndTime" },
    { 0x5FF6 , "PidTagRecipientDisplayName" },
    { 0x5FF7 , "PidTagRecipientEntryId" },
    { 0x5FFB , "PidTagRecipientTrackStatusTime" },
    { 0x5FFD , "PidTagRecipientFlags" },
    { 0x5FFF , "PidTagRecipientTrackStatus" },
    { 0x6100 , "PidTagJunkIncludeContacts" },
    { 0x6101 , "PidTagJunkThreshold" },
    { 0x6102 , "PidTagJunkPermanentlyDelete" },
    { 0x6103 , "PidTagJunkAddRecipientsToSafeSendersList" },
    { 0x6107 , "PidTagJunkPhishingEnableLinks" },
    { 0x64F0 , "PidTagMimeSkeleton" },
    { 0x65C2 , "PidTagReplyTemplateId" },
    { 0x65C6 , "PidTagSecureSubmitFlags" },
    { 0x65E0 , "PidTagSourceKey" },
    { 0x65E1 , "PidTagParentSourceKey" },
    { 0x65E2 , "PidTagChangeKey" },
    { 0x65E3 , "PidTagPredecessorChangeList" },
    { 0x65E9 , "PidTagRuleMessageState" },
    { 0x65EA , "PidTagRuleMessageUserFlags" },
    { 0x65EB , "PidTagRuleMessageProvider" },
    { 0x65EC , "PidTagRuleMessageName" },
    { 0x65ED , "PidTagRuleMessageLevel" },
    { 0x65EE , "PidTagRuleMessageProviderData" },
    { 0x65F3 , "PidTagRuleMessageSequence" },
    { 0x6619 , "PidTagUserEntryId" },
    { 0x661B , "PidTagMailboxOwnerEntryId" },
    { 0x661C , "PidTagMailboxOwnerName" },
    { 0x661D , "PidTagOutOfOfficeState" },
    { 0x6622 , "PidTagSchedulePlusFreeBusyEntryId" },
    { 0x6639 , "PidTagRights" },
    { 0x663A , "PidTagHasRules" },
    { 0x663B , "PidTagAddressBookEntryId" },
    { 0x663E , "PidTagHierarchyChangeNumber" },
    { 0x6645 , "PidTagClientActions" },
    { 0x6646 , "PidTagDamOriginalEntryId" },
    { 0x6647 , "PidTagDamBackPatched" },
    { 0x6648 , "PidTagRuleError" },
    { 0x6649 , "PidTagRuleActionType" },
    { 0x664A , "PidTagHasNamedProperties" },
    { 0x6650 , "PidTagRuleActionNumber" },
    { 0x6651 , "PidTagRuleFolderEntryId" },
    { 0x666A , "PidTagProhibitReceiveQuota" },
    { 0x666C , "PidTagInConflict" },
    { 0x666D , "PidTagMaximumSubmitMessageSize" },
    { 0x666E , "PidTagProhibitSendQuota" },
    { 0x6671 , "PidTagMemberId" },
    { 0x6672 , "PidTagMemberName" },
    { 0x6673 , "PidTagMemberRights" },
    { 0x6674 , "PidTagRuleId" },
    { 0x6675 , "PidTagRuleIds" },
    { 0x6676 , "PidTagRuleSequence" },
    { 0x6677 , "PidTagRuleState" },
    { 0x6678 , "PidTagRuleUserFlags" },
    { 0x6679 , "PidTagRuleCondition" },
    { 0x6680 , "PidTagRuleActions" },
    { 0x6681 , "PidTagRuleProvider" },
    { 0x6682 , "PidTagRuleName" },
    { 0x6683 , "PidTagRuleLevel" },
    { 0x6684 , "PidTagRuleProviderData" },
    { 0x668F , "PidTagDeletedOn" },
    { 0x66A1 , "PidTagLocaleId" },
    { 0x66C3 , "PidTagCodePageId" },
    { 0x6704 , "PidTagAddressBookManageDistributionList" },
    { 0x6705 , "PidTagSortLocaleId" },
    { 0x6709 , "PidTagLocalCommitTime" },
    { 0x670A , "PidTagLocalCommitTimeMax" },
    { 0x670B , "PidTagDeletedCountTotal" },
    { 0x670E , "PidTagFlatUrlName" },
    { 0x6740 , "PidTagSentMailSvrEID" },
    { 0x6741 , "PidTagDeferredActionMessageOriginalEntryId" },
    { 0x6748 , "PidTagFolderId" },
    { 0x6749 , "PidTagParentFolderId" },
    { 0x674A , "PidTagMid" },
    { 0x674D , "PidTagInstID" },
    { 0x674E , "PidTagInstanceNum" },
    { 0x674F , "PidTagAddressBookMessageId" },
    { 0x6793 , "MetaTagIdsetExpired" },
    { 0x6796 , "MetaTagCnsetSeen" },
    { 0x67A4 , "PidTagChangeNumber" },
    { 0x67AA , "PidTagAssociated" },
    { 0x67D2 , "MetaTagCnsetRead" },
    { 0x67DA , "MetaTagCnsetSeenFAI" },
    { 0x67E5 , "MetaTagIdsetDeleted" },
    { 0x6800 , "PidTagOfflineAddressBookName" },
    { 0x6801 , "PidTagOfflineAddressBookSequence" },
    { 0x6802 , "PidTagOfflineAddressBookContainerGuid" },
    { 0x6802 , "PidTagRwRulesStream" },
    { 0x6802 , "PidTagSenderTelephoneNumber" },
    { 0x6803 , "PidTagOfflineAddressBookMessageClass" },
    { 0x6803 , "PidTagVoiceMessageSenderName" },
    { 0x6804 , "PidTagOfflineAddressBookDistinguishedName" },
    { 0x6804 , "PidTagFaxNumberOfPages" },
    { 0x6805 , "PidTagVoiceMessageAttachmentOrder" },
    { 0x6805 , "PidTagOfflineAddressBookTruncatedProperties" },
    { 0x6806 , "PidTagCallId" },
    { 0x6820 , "PidTagReportingMessageTransferAgent" },
    { 0x6834 , "PidTagSearchFolderLastUsed" },
    { 0x683A , "PidTagSearchFolderExpiration" },
    { 0x6841 , "PidTagScheduleInfoResourceType" },
    { 0x6842 , "PidTagSearchFolderId" },
    { 0x6842 , "PidTagScheduleInfoDelegatorWantsCopy" },
    { 0x6843 , "PidTagScheduleInfoDontMailDelegates" },
    { 0x6844 , "PidTagScheduleInfoDelegateNames" },
    { 0x6844 , "PidTagSearchFolderRecreateInfo" },
    { 0x6845 , "PidTagScheduleInfoDelegateEntryIds" },
    { 0x6845 , "PidTagSearchFolderDefinition" },
    { 0x6846 , "PidTagSearchFolderStorageType" },
    { 0x6846 , "PidTagGatewayNeedsToRefresh" },
    { 0x6847 , "PidTagFreeBusyPublishStart" },
    { 0x6848 , "PidTagFreeBusyPublishEnd" },
    { 0x6849 , "PidTagFreeBusyMessageEmailAddress" },
    { 0x6849 , "PidTagWlinkType" },
    { 0x684A , "PidTagWlinkFlags" },
    { 0x684A , "PidTagScheduleInfoDelegateNamesW" },
    { 0x684B , "PidTagScheduleInfoDelegatorWantsInfo" },
    { 0x684B , "PidTagWlinkOrdinal" },
    { 0x684C , "PidTagWlinkEntryId" },
    { 0x684D , "PidTagWlinkRecordKey" },
    { 0x684E , "PidTagWlinkStoreEntryId" },
    { 0x684F , "PidTagScheduleInfoMonthsMerged" },
    { 0x684F , "PidTagWlinkFolderType" },
    { 0x6850 , "PidTagScheduleInfoFreeBusyMerged" },
    { 0x6850 , "PidTagWlinkGroupClsid" },
    { 0x6851 , "PidTagScheduleInfoMonthsTentative" },
    { 0x6851 , "PidTagWlinkGroupName" },
    { 0x6852 , "PidTagScheduleInfoFreeBusyTentative" },
    { 0x6852 , "PidTagWlinkSection" },
    { 0x6853 , "PidTagWlinkCalendarColor" },
    { 0x6853 , "PidTagScheduleInfoMonthsBusy" },
    { 0x6854 , "PidTagScheduleInfoFreeBusyBusy" },
    { 0x6854 , "PidTagWlinkAddressBookEID" },
    { 0x6855 , "PidTagScheduleInfoMonthsAway" },
    { 0x6856 , "PidTagScheduleInfoFreeBusyAway" },
    { 0x6868 , "PidTagFreeBusyRangeTimestamp" },
    { 0x6869 , "PidTagFreeBusyCountMonths" },
    { 0x686A , "PidTagScheduleInfoAppointmentTombstone" },
    { 0x686B , "PidTagDelegateFlags" },
    { 0x686C , "PidTagScheduleInfoFreeBusy" },
    { 0x686D , "PidTagScheduleInfoAutoAcceptAppointments" },
    { 0x686E , "PidTagScheduleInfoDisallowRecurringAppts" },
    { 0x686F , "PidTagScheduleInfoDisallowOverlappingAppts" },
    { 0x6890 , "PidTagWlinkClientID" },
    { 0x6891 , "PidTagWlinkAddressBookStoreEID" },
    { 0x6892 , "PidTagWlinkROGroupType" },
    { 0x7001 , "PidTagViewDescriptorBinary" },
    { 0x7002 , "PidTagViewDescriptorStrings" },
    { 0x7006 , "PidTagViewDescriptorName" },
    { 0x7007 , "PidTagViewDescriptorVersion" },
    { 0x7C06 , "PidTagRoamingDatatypes" },
    { 0x7C07 , "PidTagRoamingDictionary" },
    { 0x7C08 , "PidTagRoamingXmlStream" },
    { 0x7C24 , "PidTagOscSyncEnabled" },
    { 0x7D01 , "PidTagProcessed" },
    { 0x7FF9 , "PidTagExceptionReplaceTime" },
    { 0x7FFA , "PidTagAttachmentLinkId" },
    { 0x7FFB , "PidTagExceptionStartTime" },
    { 0x7FFC , "PidTagExceptionEndTime" },
    { 0x7FFD , "PidTagAttachmentFlags" },
    { 0x7FFE , "PidTagAttachmentHidden" },
    { 0x7FFF , "PidTagAttachmentContactPhoto" },
    { 0x8004 , "PidTagAddressBookFolderPathname" },
    { 0x8005 , "PidTagAddressBookManager" },
    { 0x8005 , "PidTagAddressBookManagerDistinguishedName" },
    { 0x8006 , "PidTagAddressBookHomeMessageDatabase" },
    { 0x8008 , "PidTagAddressBookIsMemberOfDistributionList" },
    { 0x8009 , "PidTagAddressBookMember" },
    { 0x800C , "PidTagAddressBookOwner" },
    { 0x800E , "PidTagAddressBookReports" },
    { 0x800F , "PidTagAddressBookProxyAddresses" },
    { 0x8011 , "PidTagAddressBookTargetAddress" },
    { 0x8015 , "PidTagAddressBookPublicDelegates" },
    { 0x8024 , "PidTagAddressBookOwnerBackLink" },
    { 0x802D , "PidTagAddressBookExtensionAttribute1" },
    { 0x802E , "PidTagAddressBookExtensionAttribute2" },
    { 0x802F , "PidTagAddressBookExtensionAttribute3" },
    { 0x8030 , "PidTagAddressBookExtensionAttribute4" },
    { 0x8031 , "PidTagAddressBookExtensionAttribute5" },
    { 0x8032 , "PidTagAddressBookExtensionAttribute6" },
    { 0x8033 , "PidTagAddressBookExtensionAttribute7" },
    { 0x8034 , "PidTagAddressBookExtensionAttribute8" },
    { 0x8035 , "PidTagAddressBookExtensionAttribute9" },
    { 0x8036 , "PidTagAddressBookExtensionAttribute10" },
    { 0x803C , "PidTagAddressBookObjectDistinguishedName" },
    { 0x806A , "PidTagAddressBookDeliveryContentLength" },
    { 0x8073 , "PidTagAddressBookDistributionListMemberSubmitAccepted" },
    { 0x8170 , "PidTagAddressBookNetworkAddress" },
    { 0x8C57 , "PidTagAddressBookExtensionAttribute11" },
    { 0x8C58 , "PidTagAddressBookExtensionAttribute12" },
    { 0x8C59 , "PidTagAddressBookExtensionAttribute13" },
    { 0x8C60 , "PidTagAddressBookExtensionAttribute14" },
    { 0x8C61 , "PidTagAddressBookExtensionAttribute15" },
    { 0x8C6A , "PidTagAddressBookX509Certificate" },
    { 0x8C6D , "PidTagAddressBookObjectGuid" },
    { 0x8C8E , "PidTagAddressBookPhoneticGivenName" },
    { 0x8C8F , "PidTagAddressBookPhoneticSurname" },
    { 0x8C90 , "PidTagAddressBookPhoneticDepartmentName" },
    { 0x8C91 , "PidTagAddressBookPhoneticCompanyName" },
    { 0x8C92 , "PidTagAddressBookPhoneticDisplayName" },
    { 0x8C93 , "PidTagAddressBookDisplayTypeExtended" },
    { 0x8C94 , "PidTagAddressBookHierarchicalShowInDepartments" },
    { 0x8C96 , "PidTagAddressBookRoomContainers" },
    { 0x8C97 , "PidTagAddressBookHierarchicalDepartmentMembers" },
    { 0x8C98 , "PidTagAddressBookHierarchicalRootDepartment" },
    { 0x8C99 , "PidTagAddressBookHierarchicalParentDepartment" },
    { 0x8C9A , "PidTagAddressBookHierarchicalChildDepartments" },
    { 0x8C9E , "PidTagThumbnailPhoto" },
    { 0x8CA0 , "PidTagAddressBookSeniorityIndex" },
    { 0x8CA8 , "PidTagAddressBookOrganizationalUnitRootDistinguishedName" },
    { 0x8CAC , "PidTagAddressBookSenderHintTranslations" },
    { 0x8CB5 , "PidTagAddressBookModerationEnabled" },
    { 0x8CC2 , "PidTagSpokenName" },
    { 0x8CD8 , "PidTagAddressBookAuthorizedSenders" },
    { 0x8CD9 , "PidTagAddressBookUnauthorizedSenders" },
    { 0x8CDA , "PidTagAddressBookDistributionListMemberSubmitRejected" },
    { 0x8CDB , "PidTagAddressBookDistributionListRejectMessagesFromDLMembers" },
    { 0x8CDD , "PidTagAddressBookHierarchicalIsHierarchicalGroup" },
    { 0x8CE2 , "PidTagAddressBookDistributionListMemberCount" },
    { 0x8CE3 , "PidTagAddressBookDistributionListExternalMemberCount" },
    { 0xFFFB , "PidTagAddressBookIsMaster" },
    { 0xFFFC , "PidTagAddressBookParentEntryId" },
    { 0xFFFD , "PidTagAddressBookContainerId" }
};

char* idAll2name(UINT16 id)
{
    int i;
    for (i = 0; i < sizeof(propArr)/sizeof(CFBProperty); i++)
    {
        if (id == propArr[i].id)
            return propArr[i].name;
    }
    return 0;
};


char* id2type(const char* id)
{
    // Fixed Length Properties
    if (!strcmp(id, "0002"))
        return "PtypInteger16";
    else if (!strcmp(id, "0003"))
        return "PtypInteger32";
    else if (!strcmp(id, "0004"))  //
        return "PtypFloating32";
    else if (!strcmp(id, "0005"))
        return "PtypFloating64";
    else if (!strcmp(id, "0006"))  //
        return "PtypCurrency";
    else if (!strcmp(id, "0007"))  //
        return "PtypFloatingTime";
    else if (!strcmp(id, "000A"))  //
        return "PtypErrorCode";
    else if (!strcmp(id, "000B"))
        return "PtypBoolean";
    else if (!strcmp(id, "0014"))
        return "PtypInteger64";
    else if (!strcmp(id, "0040"))
        return "PtypTime";
    else if (!strcmp(id, "0048"))  //
        return "PtypGuid";

    // Variable Length Properties
    else if (!strcmp(id, "001E"))
        return "PtypString8";
    else if (!strcmp(id, "001F"))
        return "PtypString";
    else if (!strcmp(id, "0102"))
        return "PtypBinary";
    else if (!strcmp(id, "000D"))  //
        return "PtypObject";

    // 2.1.4 Multiple-Valued Properties
    // 2.1.4.1 Fixed Length Multiple - Valued Properties

    else if (!strcmp(id, "1002"))
        return "PtypMultipleInteger16";
    else if (!strcmp(id, "1003"))
        return "PtypMultipleInteger32";
    else if (!strcmp(id, "1004"))  //
        return "PtypMultipleFloating32";
    else if (!strcmp(id, "1005"))
        return "PtypMultipleFloating64";
    else if (!strcmp(id, "1006"))  //
        return "PtypMultipleCurrency";
    else if (!strcmp(id, "1007"))  //
        return "PtypMultipleFloatingTime";
    else if (!strcmp(id, "1014"))
        return "PtypMultipleInteger64";
    else if (!strcmp(id, "1040"))
        return "PtypMultipleTime";
    else if (!strcmp(id, "1048"))  //
        return "PtypMultipleGuid";


    // 2.1.4.2 Variable Length Multiple-Valued Properties

    else if (!strcmp(id, "1102"))
        return "PtypMultipleBinary";
    else if (!strcmp(id, "101E"))
        return "PtypMultipleString8";
    else if (!strcmp(id, "101F"))
        return "PtypMultipleString";

    // End
    else
        return 0;
}


CFBProperty propMailFilterArr[] = {
    { 0x1013 , "PidTagBodyHtml" }, // Data type: PtypString, 0x001F
    { 0x1009 , "PidTagRtfCompressed" },
    { 0x0003 , "PidTagNativeBody" },   // see 0x1016
    { 0x1016 , "PidTagNativeBody" },
    { 0x1000 , "PidTagBody" },
    { 0x1009 , "PidTagRtfCompressed" },
    { 0x1042 , "PidTagInReplyToId" },
    { 0x3001 , "PidTagDisplayName" },
    { 0x5FF6 , "PidTagRecipientDisplayName" },
    { 0x3002 , "PidTagAddressType" },
    { 0x3003 , "PidTagEmailAddress" },
    { 0x39FE , "PidTagSmtpAddress" },
    { 0x0E02 , "PidTagDisplayBcc" },
    { 0x0E03 , "PidTagDisplayCc" },
    { 0x0E04 , "PidTagDisplayTo" },
    { 0x0072 , "PidTagOriginalDisplayBcc" },
    { 0x0073 , "PidTagOriginalDisplayCc" },
    { 0x0074 , "PidTagOriginalDisplayTo" },
    { 0x0C15 , "PidTagRecipientType" },
    { 0x0C1A , "PidTagSenderName" },
    { 0x0C1F , "PidTagSenderEmailAddress" },
    { 0x0E08 , "PidTagMessageSize" },
    { 0x0E12 , "PidTagMessageRecipients" },
    { 0x0E13 , "PidTagMessageAttachments" },
    { 0x0E1B , "PidTagHasAttachments" },
    { 0x0E20 , "PidTagAttachSize" },
    { 0x0037 , "PidTagSubject" },
    { 0x003D , "PidTagSubjectPrefix" },
    { 0x0049 , "PidTagOriginalSubject" },
    { 0x0E1D , "PidTagNormalizedSubject" },
    { 0x007D , "PidTagTransportMessageHeaders" },
    { 0x0E07 , "PidTagMessageFlags" },
    { 0x001A , "PidTagMessageClass" },
    { 0x3A0C , "PidTagLanguage" },
    { 0x3707 , "PidTagAttachLongFilename" },
    { 0x3703 , "PidTagAttachExtension" },
    { 0x3704 , "PidTagAttachFilename" },
    { 0x370E , "PidTagAttachMimeTag" },
    { 0x3701 , "PidTagAttachDataObject" },
    { 0x3702 , "PidTagAttachEncoding" },
    { 0x340D , "PidTagStoreSupportMask" },
    { 0x3FF1 , "PidTagMessageLocaleId" },
    { 0x3FFD , "PidTagMessageCodepage" },
    { 0x3FDE , "PidTagInternetCodepage" },
    { 0x5D01 , "PidTagSenderSmtpAddress" },
    { 0x5D02 , "PidTagSentRepresentingSmtpAddress" },
    { 0x5D05 , "PidTagReadReceiptSmtpAddress" },
    { 0x5D07 , "PidTagReceivedBySmtpAddress" },
    { 0x5D08 , "PidTagReceivedRepresentingSmtpAddress" },

};

char* idMailFilter2name(UINT16 id)
{
    int i;
    for (i = 0; i < sizeof(propMailFilterArr) / sizeof(CFBProperty); i++)
    {
        if (id == propMailFilterArr[i].id)
            return propMailFilterArr[i].name;
    }
    return 0;
};

CFBProperty propBodyFilterArr[] = {
    { 0x1013 , "PidTagBodyHtml" }, // Data type: PtypString, 0x001F
    { 0x1009 , "PidTagRtfCompressed" },
    { 0x0003 , "PidTagNativeBody" },   // see 0x1016
    { 0x1016 , "PidTagNativeBody" },
    { 0x1000 , "PidTagBody" },
    { 0x1009 , "PidTagRtfCompressed" },
    { 0x0E07 , "PidTagMessageFlags" },
    { 0x001A , "PidTagMessageClass" },
    { 0x007D , "PidTagTransportMessageHeaders" },
    { 0x340D , "PidTagStoreSupportMask" },
    { 0x3FF1 , "PidTagMessageLocaleId" },
    { 0x3FFD , "PidTagMessageCodepage" },
    { 0x3FDE , "PidTagInternetCodepage" },
    { 0x1014 , "PidTagBodyContentLocation" },
    { 0x1015 , "PidTagBodyContentId" },
};


char* idBody2name(UINT16 id)
{
    int i;
    for (i = 0; i < sizeof(propBodyFilterArr) / sizeof(CFBProperty); i++)
    {
        if (id == propBodyFilterArr[i].id)
            return propBodyFilterArr[i].name;
    }
    return 0;
};

CFBProperty propAttachmentFilterArr[] = {
    { 0x3702 , "PidTagAttachEncoding" },
    { 0x3001 , "PidTagDisplayName" },
    { 0x3701 , "PidTagAttachDataObject" },
    { 0x370E , "PidTagAttachMimeTag" },
    { 0x3704 , "PidTagAttachFilename" },
    { 0x3703 , "PidTagAttachExtension" },
    { 0x3707 , "PidTagAttachLongFilename" },
    { 0x0E07 , "PidTagMessageFlags" },  // not very important
    { 0x001A , "PidTagMessageClass" },
    { 0x3A0C , "PidTagLanguage" },
    { 0x3FF1 , "PidTagMessageLocaleId" },
    { 0x3FFD , "PidTagMessageCodepage" },
    { 0x3FDE , "PidTagInternetCodepage" },
};

char* idAttachment2name(unsigned short id)
{
    int i;
    for (i = 0; i < sizeof(propAttachmentFilterArr) / sizeof(CFBProperty); i++)
    {
        if (id == propAttachmentFilterArr[i].id)
            return propAttachmentFilterArr[i].name;
    }
    return 0;
};


bool isFixedLengthType(UINT16 type)
{
    // Fixed Length Properties
    if ((type == PtypInteger16) ||
        (type == PtypInteger32) ||
        (type == PtypFloating32) ||
        (type == PtypFloating64) ||
        (type == PtypCurrency) ||
        (type == PtypFloatingTime) ||
        (type == PtypErrorCode) ||
        (type == PtypBoolean) ||
        (type == PtypInteger64) ||
        (type == PtypTime) ||
        (type == PtypGuid))
    {
        return true;
    }
    else
        return false;
}
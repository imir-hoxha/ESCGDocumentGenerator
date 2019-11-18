using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    public class BriefingProperties
    {
        public const string PREVIOUS_SGRESPONSIBLE_KEY = "pol-previousSGResponsible";
        public const string NOT_PUBLISHED_KEY = "pol_NrOfUnpublishedDocuments";

        public Int32? Id { get; set; }
        public Guid UniqueId { get; set; }
        public Guid ListId { get; set; }
        public string ServerRelativeUrl { get; set; }
        public string Title { get; set; }
        public DateTime? StartEventDate { get; set; }
        public DateTime? EndEventDate { get; set; }
        public bool IsProtected { get; set; }
        public string Location { get; set; }
        public DateTime? RequestDate { get; set; }
        public DateTime? DraftDeadline { get; set; }
        public DateTime? FinalDeadline { get; set; }
        public DateTime? ServicesDeadline { get; set; }
        public string BriefingNumber { get; set; }
        public string[] CabMemberResponsible { get; set; }
        public string[] CabMemberResponsibleName { get; set; }
        public string[] SGResponsible { get; set; }
        public string[] SGResponsibleName { get; set; }
        public string[] OtherResponsible { get; set; }
        public string[] OtherResponsibleName { get; set; }
        public string[] UnitResponsible { get; set; }
        public string LinkToReport { get; set; }
        public bool UrgentRequest { get; set; }
        public int? NumberOfDays { get; set; }
        public bool EventCancelled { get; set; }
        public string Beneficiary { get; set; }
        public string Comments { get; set; }
        public string Category { get; set; }
        public string CategoryName { get { return Category; } set { } }
        public bool Finalized { get; set; }
        public string[] NoteTaker { get; set; }
        public string[] NoteTakerName { get; set; }
        public string Topics { get; set; }
        public string Language { get; set; }
        //public bool? NoteTakerNeeded { get; set; }
        public string[] AdditionalDocuments { get; set; }
        public string PrintedCopies { get; set; }
        public string WhoHasRequestedMeeting { get; set; }
        public string ParticipantsCommission { get; set; }
        public string ParticipantsInterlocutors { get; set; }
        public bool? HaveTheyMetBefore { get; set; }
        public string MetInWhichCapacity { get; set; }
        public DateTime? MetWhen { get; set; }
        public string PreviousCommitmentsControversies { get; set; }
        public string WhatDoWeWant { get; set; }
        public string WhatDoTheyWant { get; set; }
        //public string MeetingObjectives { get; set; }
        public string MeetingTopics { get; set; }
        public string ARESLink { get; set; }
        public string ARESReference { get; set; }
        public string Created { get; set; }
        public string CreatedBy { get; set; }
        public string Modified { get; set; }
        public string ModifiedBy { get; set; }
        public string BasePermissions { get; set; }

        public string UnitResponsibleValue
        {
            get
            {
                SPFieldMultiColumnValue unitResponsibleValue = new SPFieldMultiColumnValue();
                if (UnitResponsible != null)
                {
                    foreach (string unit in UnitResponsible)
                    {
                        unitResponsibleValue.Add(unit);
                    }
                }
                return unitResponsibleValue.ToString();
            }
            set
            {
                if (value != null)
                {
                    var unitsValue = new SPFieldMultiColumnValue(value);
                    UnitResponsible = new string[unitsValue.Count];
                    for (int i = 0; i < unitsValue.Count; i++)
                    {
                        UnitResponsible[i] = unitsValue[i];
                    }
                }
                else
                {
                    UnitResponsible = new string[] { };
                }
            }
        }

        public SPFieldUserValueCollection CabMemberResponsibleValue
        {
            get
            {
                SPFieldUserValueCollection value = new SPFieldUserValueCollection();
                foreach (string userLogin in CabMemberResponsible)
                {
                    SPUser user = SPContext.Current.Web.EnsureUser(userLogin);
                    value.Add(new SPFieldUserValue(SPContext.Current.Web, user.ID, user.LoginName));
                }
                return value;
            }
            set
            {
                if (value != null)
                {
                    int count = value.Count(u => u.User != null);
                    this.CabMemberResponsible = new string[count];
                    this.CabMemberResponsibleName = new string[count];

                    int index = 0;
                    foreach (SPFieldUserValue user in value.Where(u => u.User != null))
                    {
                        this.CabMemberResponsible[index] = user.User.LoginName;
                        this.CabMemberResponsibleName[index] = user.User.Name;
                        index++;
                    }
                }
                else
                {
                    this.CabMemberResponsible = new string[0];
                    this.CabMemberResponsibleName = new string[0];
                }
            }
        }

        public SPFieldUserValueCollection SGResponsibleValue
        {
            get
            {
                SPFieldUserValueCollection value = new SPFieldUserValueCollection();
                foreach (string userLogin in SGResponsible)
                {
                    SPUser user = SPContext.Current.Web.EnsureUser(userLogin);
                    value.Add(new SPFieldUserValue(SPContext.Current.Web, user.ID, user.LoginName));
                }
                return value;
            }
            set
            {
                if (value != null)
                {
                    int count = value.Count(u => u.User != null);
                    this.SGResponsible = new string[count];
                    this.SGResponsibleName = new string[count];

                    int index = 0;
                    foreach (SPFieldUserValue user in value.Where(u => u.User != null))
                    {
                        this.SGResponsible[index] = user.User.LoginName;
                        this.SGResponsibleName[index] = user.User.Name;
                        index++;
                    }
                }
                else
                {
                    this.SGResponsible = new string[0];
                    this.SGResponsibleName = new string[0];
                }
            }
        }

        public SPFieldUserValueCollection OtherResponsibleValue
        {
            get
            {
                SPFieldUserValueCollection value = new SPFieldUserValueCollection();
                foreach (string userLogin in OtherResponsible)
                {
                    SPUser user = SPContext.Current.Web.EnsureUser(userLogin);
                    value.Add(new SPFieldUserValue(SPContext.Current.Web, user.ID, user.LoginName));
                }
                return value;
            }
            set
            {
                if (value != null)
                {
                    int count = value.Count(u => u.User != null);
                    this.OtherResponsible = new string[count];
                    this.OtherResponsibleName = new string[count];

                    int index = 0;
                    foreach (SPFieldUserValue user in value.Where(u => u.User != null))
                    {
                        this.OtherResponsible[index] = user.User.LoginName;
                        this.OtherResponsibleName[index] = user.User.Name;
                        index++;
                    }
                }
                else
                {
                    this.OtherResponsible = new string[0];
                    this.OtherResponsibleName = new string[0];
                }
            }
        }

        public SPFieldUserValueCollection NoteTakerValue
        {
            get
            {
                SPFieldUserValueCollection value = new SPFieldUserValueCollection();
                foreach (string userLogin in NoteTaker)
                {
                    SPUser user = SPContext.Current.Web.EnsureUser(userLogin);
                    value.Add(new SPFieldUserValue(SPContext.Current.Web, user.ID, user.LoginName));
                }
                return value;
            }
            set
            {
                if (value != null)
                {
                    int count = value.Count(u => u.User != null);
                    this.NoteTaker = new string[count];
                    this.NoteTakerName = new string[count];

                    int index = 0;
                    foreach (SPFieldUserValue user in value.Where(u => u.User != null))
                    {
                        this.NoteTaker[index] = user.User.LoginName;
                        this.NoteTakerName[index] = user.User.Name;
                        index++;
                    }
                }
                else
                {
                    this.NoteTaker = new string[0];
                    this.NoteTakerName = new string[0];
                }
            }
        }

        public string NumberOfDaysValue
        {
            get
            {
                return NumberOfDays.HasValue
                    ? NumberOfDays.Value == 1 ? "1 day" : NumberOfDays.Value + " days"
                    : null;
            }
            set
            {
                NumberOfDays = value != null && value.IndexOf(' ') != -1
                    ? new Nullable<Int32>(Convert.ToInt32(value.Split(' ')[0]))
                    : null;
            }
        }

        public SPFieldMultiColumnValue AdditionalDocumentsValue
        {
            get
            {
                var value = new SPFieldMultiColumnValue();
                if (AdditionalDocuments != null)
                    foreach (var doc in AdditionalDocuments)
                    {
                        value.Add(doc);
                    }
                return value;
            }
            set
            {
                AdditionalDocuments = value.ColumnValues.ToArray();
            }
        }

        public int NotPublished { get; set; }

        public BriefingProperties() { }

        public BriefingProperties(SPWeb web, Guid listId, int id) : this(web, web.Lists[listId].GetItemById(id)) { }

        public BriefingProperties(SPWeb web, SPListItem briefingItem)
        {
            SPList list = briefingItem.ParentList;
            Id = briefingItem.ID;
            ListId = list.ID;
            UniqueId = briefingItem.UniqueId;
            ServerRelativeUrl = briefingItem.Folder.ServerRelativeUrl;

            Title = (string)briefingItem[BriefingFields.BriefingTitleId];
            StartEventDate = (DateTime?)briefingItem[BriefingFields.StartEventDateId];
            EndEventDate = (DateTime?)briefingItem[BriefingFields.EndEventDateId];
            IsProtected = briefingItem.Fields.ContainsField(BriefingFields.IsProtectedName) && briefingItem[BriefingFields.IsProtectedName] != null ? (bool)briefingItem[BriefingFields.IsProtectedName] : false;
            Location = (string)briefingItem[BriefingFields.LocationId];
            RequestDate = (DateTime?)briefingItem[BriefingFields.RequestDateName];
            DraftDeadline = (DateTime?)briefingItem[BriefingFields.DraftDeadlineId];
            FinalDeadline = (DateTime?)briefingItem[BriefingFields.FinalDeadlineId];
            ServicesDeadline = (DateTime?)briefingItem[BriefingFields.ServicesDeadlineId];
            BriefingNumber = (string)briefingItem[BriefingFields.BriefingNumberId];
            var cabMemberResponsible = briefingItem[BriefingFields.CabMemberResponsibleId] as SPFieldUserValueCollection;
            CabMemberResponsibleValue = cabMemberResponsible;
            var sgResponsible = briefingItem[BriefingFields.SGResponsibleId] as SPFieldUserValueCollection;
            SGResponsibleValue = sgResponsible;
            var otherResponsible = briefingItem[BriefingFields.OtherResponsibleId] as SPFieldUserValueCollection;
            OtherResponsibleValue = otherResponsible;
            UnitResponsibleValue = briefingItem[BriefingFields.UnitResponsibleId] as string;
            LinkToReport = briefingItem[BriefingFields.LinkToReportId] as string;
            UrgentRequest = briefingItem[BriefingFields.UrgentRequestId] != null ? (bool)briefingItem[BriefingFields.UrgentRequestId] : false;
            NumberOfDaysValue = briefingItem[BriefingFields.BriefingUrgencyId] as string;
            bool? eventCancelled = (bool?)briefingItem[BriefingFields.CancelledId];
            EventCancelled = eventCancelled.HasValue ? eventCancelled.Value : false;
            Beneficiary = briefingItem[BriefingFields.BeneficiaryId] as string;
            Comments = briefingItem[BriefingFields.CommentsId] as string;
            SPFolder parentFolder = briefingItem.Folder.ParentFolder;
            Category = parentFolder.Item != null && parentFolder.Item.ContentType.Name == ContentTypes.CategoryContentTypeName ? parentFolder.Name : null;
            bool? finalized = (bool?)briefingItem[BriefingFields.FinalizedName];
            Finalized = finalized.HasValue ? finalized.Value : false;
            var noteTaker = briefingItem[BriefingFields.NoteTakerId] as SPFieldUserValueCollection;
            NoteTakerValue = noteTaker;

            Language = briefingItem.Fields.Contains(BriefingFields.Language) ? briefingItem[BriefingFields.Language] as string : null;
            AdditionalDocumentsValue = briefingItem.Fields.Contains(BriefingFields.AdditionalDocuments) ? new SPFieldMultiColumnValue((string)briefingItem[BriefingFields.AdditionalDocuments]) : new SPFieldMultiColumnValue();
            PrintedCopies = briefingItem.Fields.Contains(BriefingFields.PrintedCopies) ? (string)briefingItem[BriefingFields.PrintedCopies] : null;
            WhoHasRequestedMeeting = briefingItem.Fields.Contains(BriefingFields.WhoHasRequestedMeeting) ? (string)briefingItem[BriefingFields.WhoHasRequestedMeeting] : null;
            ParticipantsCommission = briefingItem.Fields.Contains(BriefingFields.ParticipantsCommission) ? (string)briefingItem[BriefingFields.ParticipantsCommission] : null;
            ParticipantsInterlocutors = briefingItem.Fields.Contains(BriefingFields.ParticipantsInterlocutors) ? (string)briefingItem[BriefingFields.ParticipantsInterlocutors] : null;
            HaveTheyMetBefore = briefingItem.Fields.Contains(BriefingFields.HaveTheyMetBefore) ? (bool?)briefingItem[BriefingFields.HaveTheyMetBefore] : null;
            MetInWhichCapacity = briefingItem.Fields.Contains(BriefingFields.MetInWhichCapacity) ? (string)briefingItem[BriefingFields.MetInWhichCapacity] : null;
            MetWhen = briefingItem.Fields.Contains(BriefingFields.MetWhen) ? (DateTime?)briefingItem[BriefingFields.MetWhen] : null;
            PreviousCommitmentsControversies = briefingItem.Fields.Contains(BriefingFields.PreviousCommitmentsControversies) ? (string)briefingItem[BriefingFields.PreviousCommitmentsControversies] : null;
            WhatDoTheyWant = briefingItem.Fields.Contains(BriefingFields.WhatDoTheyWant) ? (string)briefingItem[BriefingFields.WhatDoTheyWant] : null;
            WhatDoWeWant = briefingItem.Fields.Contains(BriefingFields.WhatDoWeWant) ? (string)briefingItem[BriefingFields.WhatDoWeWant] : null;
            MeetingTopics = briefingItem.Fields.Contains(BriefingFields.MeetingTopics) ? (string)briefingItem[BriefingFields.MeetingTopics] : null;

            if (briefingItem.Fields.Contains(BriefingFields.HermesLinkId) && briefingItem.Fields.Contains(BriefingFields.HermesReferenceId))
            {
                ARESLink = (string)briefingItem[BriefingFields.HermesLinkId];
                ARESReference = (string)briefingItem[BriefingFields.HermesReferenceId];
            }
            Created = ((DateTime)briefingItem[SPBuiltInFieldId.Created]).ToDateTimeString();
            CreatedBy = new SPFieldUserValue(web, (string)briefingItem[SPBuiltInFieldId.Author]).LookupValue;
            Modified = ((DateTime)briefingItem[SPBuiltInFieldId.Modified]).ToDateTimeString();
            ModifiedBy = new SPFieldUserValue(web, (string)briefingItem[SPBuiltInFieldId.Editor]).LookupValue;
            BasePermissions = briefingItem.EffectiveBasePermissions.ToString();
            NotPublished = briefingItem.Properties[NOT_PUBLISHED_KEY] != null ? Convert.ToInt32(briefingItem.Properties[NOT_PUBLISHED_KEY]) : 0;
        }

        public static string GetBriefingNumberForId(SPWeb web, Guid listId, int id)
        {
            SPList list = web.Lists[listId];
            SPListItem briefingItem = list.GetItemById(id);
            return briefingItem[BriefingFields.BriefingNumberId] as string;
        }

        public void Save(SPWeb web, bool addTemplates = true)
        {
            SPListItem briefingItem = null;
            SPList list = web.Lists[ListId];
            string folderName = FileUtil.CleanForUrlAndFileNameUse(Title);

            if (Id.HasValue)
            {
                briefingItem = list.GetItemById(Id.Value);
                briefingItem["FileLeafRef"] = folderName;
                BriefingNumber = (string)briefingItem[BriefingFields.BriefingNumberId];
            }
            else
            {
                if (string.IsNullOrEmpty(Category))
                {
                    briefingItem = list.AddItem(list.RootFolder.Url, SPFileSystemObjectType.Folder, folderName);
                }
                else
                {
                    string folderUrl = SPUrlUtility.CombineUrl(list.RootFolder.Url, Category);
                    if (web.GetFolder(folderUrl).Exists)
                        briefingItem = list.AddItem(folderUrl, SPFileSystemObjectType.Folder, folderName);
                    else
                        briefingItem = list.AddItem(list.RootFolder.Url, SPFileSystemObjectType.Folder, folderName);
                }

                SPContentType ctype = list.ContentTypes[list.ContentTypes.BestMatch(ContentTypes.BriefingContentTypeId)];
                briefingItem[SPBuiltInFieldId.ContentTypeId] = ctype.Id;
                briefingItem[SPBuiltInFieldId.ContentType] = ContentTypes.BriefingContentTypeName;
                if (string.IsNullOrEmpty(BriefingNumber))
                    BriefingNumber = Config.System.CurrentSystem == SystemType.Poline ? BriefingNumbers.GetNextPolineNumber(StartEventDate.Value.Year) : BriefingNumbers.GetNextNumber(web, StartEventDate.Value);
                briefingItem[BriefingFields.BriefingNumberId] = BriefingNumber;
            }

            briefingItem[BriefingFields.BriefingTitleId] = Title;
            briefingItem[BriefingFields.StartEventDateId] = StartEventDate;
            briefingItem[BriefingFields.EndEventDateId] = EndEventDate;
            if (briefingItem.Fields.ContainsField(BriefingFields.IsProtectedName))
                briefingItem[BriefingFields.IsProtectedName] = IsProtected;
            briefingItem[BriefingFields.LocationId] = Location;
            briefingItem[BriefingFields.MainTopicsId] = Topics;
            briefingItem[BriefingFields.RequestDateName] = RequestDate;
            briefingItem[BriefingFields.DraftDeadlineId] = DraftDeadline;
            briefingItem[BriefingFields.FinalDeadlineId] = FinalDeadline;
            briefingItem[BriefingFields.ServicesDeadlineId] = ServicesDeadline;
            briefingItem[BriefingFields.CabMemberResponsibleId] = CabMemberResponsibleValue;
            SPFieldUserValueCollection PrevSGResponsibleValue = briefingItem[BriefingFields.SGResponsibleId] as SPFieldUserValueCollection ?? new SPFieldUserValueCollection();
            if (SGResponsibleValue.Count != PrevSGResponsibleValue.Count ||
                SGResponsibleValue.Except(PrevSGResponsibleValue, new SPFieldUserValueComparer()).Any())
            {
                briefingItem.Properties[PREVIOUS_SGRESPONSIBLE_KEY] = string.Join(",", PrevSGResponsibleValue.Select(user => user.User.LoginName).ToArray());
            }
            briefingItem[BriefingFields.SGResponsibleId] = SGResponsibleValue;
            briefingItem[BriefingFields.OtherResponsibleId] = OtherResponsibleValue;
            briefingItem[BriefingFields.UnitResponsibleId] = UnitResponsibleValue;
            briefingItem[BriefingFields.LinkToReportId] = LinkToReport;
            briefingItem[BriefingFields.UrgentRequestId] = UrgentRequest;
            briefingItem[BriefingFields.BriefingUrgencyId] = NumberOfDaysValue;
            briefingItem[BriefingFields.CancelledId] = EventCancelled;
            briefingItem[BriefingFields.BeneficiaryId] = Beneficiary;
            briefingItem[BriefingFields.CommentsId] = Comments;
            briefingItem[BriefingFields.FinalizedName] = Finalized;
            briefingItem[BriefingFields.NoteTakerId] = NoteTakerValue;

            if (briefingItem.Fields.Contains(BriefingFields.Language)) briefingItem[BriefingFields.Language] = Language;
            if (briefingItem.Fields.Contains(BriefingFields.AdditionalDocuments)) briefingItem[BriefingFields.AdditionalDocuments] = AdditionalDocumentsValue.ToString();
            if (briefingItem.Fields.Contains(BriefingFields.PrintedCopies)) briefingItem[BriefingFields.PrintedCopies] = PrintedCopies;
            if (briefingItem.Fields.Contains(BriefingFields.WhoHasRequestedMeeting)) briefingItem[BriefingFields.WhoHasRequestedMeeting] = WhoHasRequestedMeeting;
            if (briefingItem.Fields.Contains(BriefingFields.ParticipantsCommission)) briefingItem[BriefingFields.ParticipantsCommission] = ParticipantsCommission;
            if (briefingItem.Fields.Contains(BriefingFields.ParticipantsInterlocutors)) briefingItem[BriefingFields.ParticipantsInterlocutors] = ParticipantsInterlocutors;
            if (briefingItem.Fields.Contains(BriefingFields.HaveTheyMetBefore)) briefingItem[BriefingFields.HaveTheyMetBefore] = HaveTheyMetBefore;
            if (briefingItem.Fields.Contains(BriefingFields.MetInWhichCapacity)) briefingItem[BriefingFields.MetInWhichCapacity] = MetInWhichCapacity;
            if (briefingItem.Fields.Contains(BriefingFields.MetWhen)) briefingItem[BriefingFields.MetWhen] = MetWhen;
            if (briefingItem.Fields.Contains(BriefingFields.PreviousCommitmentsControversies)) briefingItem[BriefingFields.PreviousCommitmentsControversies] = PreviousCommitmentsControversies;
            if (briefingItem.Fields.Contains(BriefingFields.WhatDoTheyWant)) briefingItem[BriefingFields.WhatDoTheyWant] = WhatDoTheyWant;
            if (briefingItem.Fields.Contains(BriefingFields.WhatDoWeWant)) briefingItem[BriefingFields.WhatDoWeWant] = WhatDoWeWant;
            if (briefingItem.Fields.Contains(BriefingFields.MeetingTopics)) briefingItem[BriefingFields.MeetingTopics] = MeetingTopics;

            if (briefingItem.Fields.Contains(BriefingFields.HermesLinkId) && briefingItem.Fields.Contains(BriefingFields.HermesReferenceId))
            {
                ARESLink = (string)briefingItem[BriefingFields.HermesLinkId];
                ARESReference = (string)briefingItem[BriefingFields.HermesReferenceId];
            }

            briefingItem.Update();

            ServerRelativeUrl = briefingItem.Folder.ServerRelativeUrl;

            SPUserToken userToken = SPContext.Current.Site.UserToken;
            Guid siteId = web.Site.ID;
            Guid webId = web.ID;
            int itemId = briefingItem.ID;

            if (Id.HasValue)
            {
                if (briefingItem.Folder.ParentFolder.Name != Category)
                {
                    SPFolder folder = briefingItem.Folder;
                    if (string.IsNullOrEmpty(Category))
                    {
                        ServerRelativeUrl = SPUrlUtility.CombineUrl(list.RootFolder.ServerRelativeUrl, folderName);
                    }
                    else
                    {
                        ServerRelativeUrl = SPUrlUtility.CombineUrl(SPUrlUtility.CombineUrl(list.RootFolder.ServerRelativeUrl, Category), folderName);
                    }
                    if (briefingItem.Folder.ServerRelativeUrl != ServerRelativeUrl)
                    {
                        folder.MoveTo(ServerRelativeUrl);
                        briefingItem = list.GetItemById(Id.Value);
                    }
                }
            }
            else
            {
                HostingEnvironment.QueueBackgroundWorkItem(token => copyTemplatesAsync(siteId, webId, ListId, itemId, userToken, addTemplates, IsProtected));
            }

            string bn = Id.HasValue ? BriefingNumber : web.CurrentUser.LoginName.Substring(web.CurrentUser.LoginName.LastIndexOf("|") + 1) + "-new-briefing";
            HostingEnvironment.QueueBackgroundWorkItem(token => saveAdditionalDocumentsAsync(siteId, webId, userToken, ListId, itemId, bn, AdditionalDocuments));

            Id = briefingItem.ID;
            UniqueId = briefingItem.Folder.UniqueId;
            Modified = ((DateTime)briefingItem[SPBuiltInFieldId.Modified]).ToString();
            ModifiedBy = new SPFieldUserValue(web, (string)briefingItem[SPBuiltInFieldId.Editor]).LookupValue;
            BasePermissions = briefingItem.EffectiveBasePermissions.ToString();

            HostingEnvironment.QueueBackgroundWorkItem(token => ensureBriefingProtectedAsync(siteId, webId, ListId, itemId, IsProtected));
        }

        private static async Task saveAdditionalDocumentsAsync(Guid siteId, Guid webId, SPUserToken userToken, Guid listId, int itemId, string briefingNumber, string[] additionalDocuments)
        {
            await Task.Run(() =>
            {
                using (SPSite site = new SPSite(siteId, userToken))
                using (SPWeb web = site.OpenWeb(webId))
                {
                    SPList list = web.Lists[listId];
                    SPListItem briefingItem = list.GetItemById(itemId);
                    saveAdditionalDocuments(web, briefingItem, briefingNumber, additionalDocuments);
                }
            });
        }

        private static void saveAdditionalDocuments(SPWeb web, SPListItem briefingItem, string briefingNumber, string[] additionalDocuments)
        {
            SPFolder additionalDocumentsFolder = web.GetFolder(SPUrlUtility.CombineUrl(briefingItem.Folder.Url, "/Additional Documents"));

            if (additionalDocumentsFolder.Exists)
            {
                if (additionalDocuments != null)
                {
                    foreach (SPFile file in additionalDocumentsFolder.Files.OfType<SPFile>().ToList())
                    {
                        if (!additionalDocuments.Contains(file.Name))
                        {
                            file.Recycle();
                        }
                    }
                }
            }
            else
            {
                additionalDocumentsFolder = briefingItem.Folder.SubFolders.Add("Additional Documents");
            }

            SPSecurity.RunWithElevatedPrivileges(() =>
            {
                using (var ctx = new SPFolderContext())
                {
                    var tempDocuments = ctx.BriefingInputDocuments
                        .Where(doc => doc.BriefingNumber == briefingNumber)
                        .Where(doc => additionalDocuments.Contains(doc.Filename));
                    foreach (var file in tempDocuments)
                    {
                        additionalDocumentsFolder.Files.Add(file.Filename, file.Filecontent, true);
                    }
                    ctx.ClearBriefingInputDocuments(briefingNumber);
                }
            });
        }

        public static string CopyBriefing(SPWeb web, int itemId, Guid listId, Guid destWebId, Guid destListId, bool addSuffix)
        {
            string currentUser = string.Format("{0};#{1}", web.CurrentUser.ID, web.CurrentUser.LoginName);

            BriefingProperties briefing = new BriefingProperties(web, listId, itemId);
            ContributionEvent contributionEvent = new ContributionEvent(web, briefing.BriefingNumber, briefing.UniqueId, briefing.ListId);
            contributionEvent.Id = null;
            foreach (var req in contributionEvent.Requests)
            {
                req.Id = null;
            }
            contributionEvent.StartEventDate = contributionEvent.StartEventDate ?? briefing.StartEventDate;
            contributionEvent.EndEventDate = contributionEvent.EndEventDate ?? briefing.EndEventDate;
            string eventServerRelativeUrl = contributionEvent.ServerRelativeUrl;

            briefing.ListId = destListId;
            briefing.Id = null;

            using (SPWeb destWeb = web.Site.OpenWeb(destWebId))
            {
                destWeb.AllowUnsafeUpdates = true;
                SPList list = destWeb.Lists[briefing.ListId];
                string folderName = FileUtil.CleanForUrlAndFileNameUse(briefing.Title);
                string destBriefingServerRelativeUrl = string.IsNullOrEmpty(briefing.Category)
                    ? SPUrlUtility.CombineUrl(list.RootFolder.ServerRelativeUrl, folderName)
                    : SPUrlUtility.CombineUrl(SPUrlUtility.CombineUrl(list.RootFolder.ServerRelativeUrl, briefing.Category), folderName);
                if (addSuffix)
                {
                    if (destWeb.GetFolder(destBriefingServerRelativeUrl).Exists)
                    {
                        for (int i = 2; i < 100; i++)
                        {
                            if (destWeb.GetFolder($"{destBriefingServerRelativeUrl} {i}").Exists)
                            {
                                continue;
                            }
                            else
                            {
                                briefing.Title = briefing.Title + " " + i;
                                break;
                            }
                        }
                    }
                }
                else if (destWeb.GetFolder(destBriefingServerRelativeUrl).Exists)
                {
                    return "Another briefing with the same title exists in the destination, would you like to add a suffix number?";
                }
                string briefingNumber = briefing.BriefingNumber;
                briefing.Save(destWeb, false);
                briefing.BriefingNumber = briefingNumber;
                SPListItem item = list.GetItemById(briefing.Id.Value);
                item[BriefingFields.BriefingNumberId] = briefingNumber;
                item.Update();

                destWeb.AllowUnsafeUpdates = true;
                contributionEvent.Save(destWeb, briefing, currentUser, null);
                copyBriefingFiles(
                    web.GetFolder(eventServerRelativeUrl),
                    destWeb.GetFolder(contributionEvent.ServerRelativeUrl)
                );

                SPListItem srcItem = web.Lists[listId].GetItemById(itemId);
                copyBriefingFiles(srcItem.Folder, item.Folder);
            }

            return null;
        }

        public static void Recycle(SPWeb web, int itemId, Guid listId)
        {
            SPListItem briefingItem = web.Lists[listId].GetItemById(itemId);
            web.AllowUnsafeUpdates = true;
            briefingItem.Recycle();
        }

        private static void ensureBriefingProtectedAsync(Guid siteId, Guid webId, Guid listId, int itemId, bool isProtected)
        {
            SPSecurity.RunWithElevatedPrivileges(() =>
            {
                using (SPSite site = new SPSite(siteId))
                using (SPWeb web = site.OpenWeb(webId))
                {
                    SPList list = web.Lists[listId];
                    SPListItem briefingItem = list.GetItemById(itemId);

                    ensureBriefingProtected(web, briefingItem, isProtected);
                }
            });
        }

        private static void ensureBriefingProtected(SPWeb web, SPListItem briefingItem, bool isProtected)
        {
            using (var scope = new SPMonitoredScope("Ensure briefing is protected", TraceSeverity.Verbose, new SPSqlQueryCounter()))
            {
                if (isProtected)
                {
                    if (!briefingItem.HasUniqueRoleAssignments)
                    {
                        web.AllowUnsafeUpdates = true;
                        secureItem(web, briefingItem);
                        secureFolder(web, briefingItem.Folder);
                    }
                }
                else if (briefingItem.HasUniqueRoleAssignments)
                {
                    web.AllowUnsafeUpdates = true;
                    inheritItemPermissions(briefingItem);
                    inheritFolderPermissions(briefingItem.Folder);
                }
            }
        }

        private static void secureFolder(SPWeb web, SPFolder folder)
        {
            foreach (SPFile doc in folder.Files)
            {
                if (!doc.Name.EndsWith(".html"))
                    secureItem(web, doc.Item);
            }
            foreach (SPFolder subFolder in folder.SubFolders)
            {
                secureFolder(web, subFolder);
            }
        }

        private static void inheritFolderPermissions(SPFolder folder)
        {
            foreach (SPFile doc in folder.Files)
            {
                if (!doc.Name.EndsWith(".html"))
                    inheritItemPermissions(doc.Item);
            }
            foreach (SPFolder subFolder in folder.SubFolders)
            {
                inheritFolderPermissions(subFolder);
            }
        }

        private static void secureItem(SPWeb web, SPListItem item)
        {
            if (Config.System.CurrentSystem == SystemType.Poline)
            {
                if (item.ContentType.Name == "Poline Document" && item.Fields.ContainsField(BriefingFields.IsProtectedName))
                {
                    if (item["This document is protected"] == null || (bool)item["This document is protected"] == false)
                    {
                        item["This document is protected"] = true;
                        item.Update();
                    }
                }
            }
            else
            {
                if (item.ContentType.Name == "Document" && item.Fields.ContainsField(BriefingFields.IsProtectedName))
                {
                    if (item[BriefingFields.IsProtectedName] == null || (bool)item[BriefingFields.IsProtectedName] == false)
                    {
                        item[BriefingFields.IsProtectedName] = true;
                        item.Update();
                    }
                }
            }
            item.BreakRoleInheritance(true);
            var protectedBriefingGroups = Config.System.GetProtectedBriefingGroups(web);
            for (int i = 0; i < item.RoleAssignments.Count; i++)
            {
                var assignment = item.RoleAssignments[i];
                if (!protectedBriefingGroups.Contains(assignment.Member.Name))
                {
                    assignment.RoleDefinitionBindings.RemoveAll();
                    assignment.Update();
                }
            }
        }

        private static void inheritItemPermissions(SPListItem item)
        {
            if (Config.System.CurrentSystem == SystemType.Poline)
            {
                if (item.ContentType.Name == "Poline Document" && item.Fields.ContainsField(BriefingFields.IsProtectedName))
                {
                    if (item["This document is protected"] != null && (bool)item["This document is protected"] == true)
                    {
                        item["This document is protected"] = false;
                        item.Update();
                    }
                }
            }
            else
            {
                if (item.ContentType.Name == "Document" && item.Fields.ContainsField(BriefingFields.IsProtectedName))
                {
                    if (item[BriefingFields.IsProtectedName] != null && (bool)item[BriefingFields.IsProtectedName] == true)
                    {
                        item[BriefingFields.IsProtectedName] = false;
                        item.Update();
                    }
                }
            }
            item.ResetRoleInheritance();
        }

        protected static async Task copyTemplatesAsync(Guid siteId, Guid webId, Guid listId, int itemId, SPUserToken userToken, bool addTemplates, bool isProtected)
        {
            await Task.Run(() =>
            {
                using (SPSite qSite = new SPSite(siteId, userToken))
                using (SPWeb qWeb = qSite.OpenWeb(webId))
                {
                    SPList qList = qWeb.Lists[listId];
                    SPListItem qBriefingItem = qList.GetItemById(itemId);

                    string folderUrl = Config.System.RootWebBriefingTemplatesFolder;
                    if (addTemplates && folderUrl != null)
                    {
                        SPFolder bfTemplateFolder = qWeb.ParentWeb.GetFolder(SPUrlUtility.CombineUrl(qWeb.ParentWeb.Url, folderUrl));
                        copyTemplates(bfTemplateFolder, qBriefingItem.Folder, isProtected);
                    }

                    string libraryUrl = Config.System.CurrentWebBriefingTemplatesLibrary;
                    if (libraryUrl != null)
                    {
                        SPFolder briefingFolder = qBriefingItem.Folder;
                        if (addTemplates)
                        {
                            SPList bfTemplateLibrary = qWeb.GetList(SPUrlUtility.CombineUrl(qWeb.Url, libraryUrl));
                            copyTemplates(bfTemplateLibrary.RootFolder, briefingFolder, isProtected);
                        }

                        List<SPContentType> contenttypes = new List<SPContentType>();
                        contenttypes.Add(qBriefingItem.ParentList.ContentTypes["Document"]);
                        briefingFolder.UniqueContentTypeOrder = contenttypes;
                        briefingFolder.Update();
                    }
                }
            });
        }

        private static void copyTemplates(SPFolder template, SPFolder dest, bool isProtected)
        {
            foreach (SPFile f in template.Files)
            {
                Hashtable fileProperties = new Hashtable();
                fileProperties["vti_Title"] = Regex.Replace(f.Name, "[0-9]+\\s", "");

                int intOutNr;
                string strNumber = f.Name.Substring(0, 2);
                if (int.TryParse(strNumber, out intOutNr))
                {
                    fileProperties["DocumentNumber"] = strNumber;
                }

                var templatefile = f.OpenBinaryStream(SPOpenBinaryOptions.SkipVirusScan);
                var file = dest.Files.Add(SPUrlUtility.CombineUrl(dest.ServerRelativeUrl, f.Name), templatefile, fileProperties, true);
                if (isProtected)
                {
                    file.Item[BriefingFields.DocumentIsProtectedName] = isProtected;
                    file.Item.UpdateOverwriteVersion();
                    SPUtil.SecureBriefing(file.Item);
                }
                if (file.CheckOutType != SPFile.SPCheckOutType.None)
                {
                    file.CheckIn("");
                }
            }

            foreach (SPFolder folder in template.SubFolders)
            {
                if (folder.Item != null && folder.Name != "Forms")
                {
                    SPFolder fld = dest.SubFolders.Add(folder.Name);
                    SPListItem folderItem = fld.Item;
                    folderItem["ExemptPublishAll"] = folder.Item["ExemptPublishAll"];
                    folderItem.Update();

                    copyTemplates(folder, fld, isProtected);
                }
            }
        }

        private static void copyBriefingFiles(SPFolder sourceFolder, SPFolder destinationFolder)
        {
            foreach (SPFile f in sourceFolder.Files)
            {
                if (f.Exists && !f.Name.EndsWith(".aspx", StringComparison.InvariantCultureIgnoreCase))
                {
                    var templatefile = f.OpenBinaryStream(SPOpenBinaryOptions.SkipVirusScan);
                    var file = destinationFolder.Files.Add(SPUrlUtility.CombineUrl(destinationFolder.ServerRelativeUrl, f.Name), templatefile, true);
                    if (file.CheckOutType != SPFile.SPCheckOutType.None)
                    {
                        file.CheckIn("");
                    }
                }
            }

            foreach (SPFolder folder in sourceFolder.SubFolders)
            {
                if (folder.Item != null && folder.Name != "Forms")
                {
                    SPFolder fld = destinationFolder.SubFolders.Add(folder.Name);

                    copyBriefingFiles(folder, fld);
                }
            }
        }
    }
}

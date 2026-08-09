namespace AddUsers.Services
{
    /// <summary>A SharePoint site discovered via search.</summary>
    public sealed class SiteInfo
    {
        public string Title { get; set; }
        public string Url { get; set; }

        public override string ToString()
        {
            return string.IsNullOrEmpty(Title) ? Url : Title;
        }
    }

    /// <summary>A SharePoint site group.</summary>
    public sealed class GroupInfo
    {
        public int Id { get; set; }
        public string Title { get; set; }

        public override string ToString()
        {
            return Title;
        }
    }

    /// <summary>Outcome of attempting to add a single user to a group.</summary>
    public sealed class AddUserOutcome
    {
        public string Email { get; set; }
        public bool Success { get; set; }
        public string Detail { get; set; }
    }
}

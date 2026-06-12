// SearchDocuments.tsx
import * as React from "react";
import { useState, useMemo, useEffect, useRef } from "react";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/search";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { TooltipHost, IconButton, Callout, DirectionalHint } from "@fluentui/react";
import DOMPurify from "dompurify";
import {
  TextField,
  PrimaryButton,
  DefaultButton,
  ChoiceGroup,
  IChoiceGroupOption,
  Checkbox,
  Stack,
  Spinner,
  Persona,
  PersonaSize,
  MessageBar,
  MessageBarType,
} from "@fluentui/react";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { ResponseType } from "@microsoft/microsoft-graph-client";
import { LivePersona } from "@pnp/spfx-controls-react/lib/LivePersona";

interface ISearchDocumentsProps {
  context: any;
  msGraphClientFactory: any;
  variant?: "SearchBar" | "SearchWithResult";
}

/** 🔧 CONFIG */
const SITE_URL = "https://vaughnconstruction.sharepoint.com";
const DOCUMENT_LIBRARY_ID = "8177736d-9faa-49cc-82b3-4ff5a93fa02e";

const WEEKLY_WORDS_LIST_ID = "b2b9bc28-8e97-4716-9b88-3c456f8d2b6a";
const LESSONS_LEARNED_LIST_ID = "42F972A2-3D9D-4DBE-86C8-32EBAF4ACFDB";

const WEEKLY_WORDS_SITE_URL = `${SITE_URL}/news`;
const LESSONS_LEARNED_SITE_URL = `${SITE_URL}/ll`;

const IT_SITE_URL = `${SITE_URL}/sites/ITSite`;
const IT_BLOGS_LIBRARY_ID = "7dd1ac79-1372-41a6-972f-0e6eca0d9f0e";
const IT_BLOGS_PATH_CONTAINS = "/sites/ITSite/SitePages";

const COMPANY_CAL_SITE_URL = `${SITE_URL}/sites/OPSSite`;
const COMPANY_CAL_LIST_ID = "3c14f133-7cb6-4582-ac0a-4e5c401a5952";

const CORP_RESP_SITE_URL = `${SITE_URL}/sites/opsite`;
const CORP_RESP_LIST_ID = "8c97c9e1-ecaf-4fcb-93a9-4016509d185c";

const MARKETING_CAL_SITE_URL = `${SITE_URL}/Sites/Marketing/`;
const MARKETING_CAL_LIST_ID = "9be28777-0981-4627-a366-779f280711ef";

const LOCATIONS_SITE_URL = `${SITE_URL}/`;
const LOCATIONS_LIST_ID = "450dda32-a7e7-4439-8779-e1cc1523e7fd";

const PEOPLE_SOURCE_ID = "B09A7990-05EA-4AF9-81EF-EDFAB16C4E31";
const RESULTS_PAGE_URL =
  "https://vaughnconstruction.sharepoint.com/sitepages/Custom-Search-Results.aspx";

type Mode = "exact" | "all" | "any";

const fields = ["Title", "Filename", "Content", "Author", "Tags"];

const escapeKql = (text: string) =>
  text.replace(/(["\\()])/g, "\\$1");

const buildKql = (text: string, mode: Mode, searchFields: string[]) => {
  const escapedText = escapeKql(text.trim());
  const words = escapedText.split(/\s+/);
  if (mode === "exact") {
    const phrase = `"${escapedText}"`;
    return searchFields.map((f) => `${f}:${phrase}`).join(" OR ");
  }
  if (mode === "all") {
    const perWord = words.map((w) =>
      searchFields.map((f) => `${f}:${w}`).join(" OR ")
    );
    return perWord.join(" AND ");
  }
  const clauses: string[] = [];
  for (const w of words) {
    for (const f of searchFields) {
      clauses.push(`${f}:${w}`);
    }
  }
  return clauses.join(" OR ");
};

const buildGraphQueryString = (text: string, mode: Mode) => {
  const trimmed = text.trim();
  const words = trimmed.split(/\s+/).filter(Boolean);
  if (!trimmed) return "";
  if (mode === "exact") return `"${trimmed}"`;
  if (mode === "all") return trimmed;
  if (words.length === 0) return trimmed;
  return words.join(" OR ");
};

const highlight = (text: string, words: string[]) => {
  if (!text) return "";
  let newText = text;
  words.forEach((w) => {
    if (!w) return;
    const reg = new RegExp(`(${w})`, "gi");
    newText = newText.replace(reg, `<mark style="background:yellow">$1</mark>`);
  });
  return newText;
};

const normalizeGuid = (g?: string) =>
  (g || "")
    .toLowerCase()
    .replace(/[{}]/g, "")
    .replace(/-/g, "");

const normalizeUpn = (value?: string) => {
  const v = (value || "").trim();
  if (!v) return "";
  const pipe = v.lastIndexOf("|");
  if (pipe >= 0 && pipe < v.length - 1) return v.substring(pipe + 1).toLowerCase();
  return v.toLowerCase();
};

const isGuestPerson = (p: Partial<ISearchResult>) => {
  const account = (p.AccountName || "").toLowerCase();
  const sip = (p.SipAddress || "").toLowerCase();
  const email = (p.WorkEmail || "").toLowerCase();
  if (!account && !sip && !email) return true;
  const hay = `${account} ${sip} ${email}`;
  if (hay.includes("#ext#")) return true;
  if (hay.includes("urn:spo:guest")) return true;
  return false;
};

const peopleMatchesQuery = (
  p: Partial<ISearchResult>,
  queryText: string,
  mode: Mode
) => {
  const q = (queryText || "").trim().toLowerCase();
  if (!q) return true;
  const hay = [
    p.PreferredName,
    p.Title,
    p.JobTitle,
    p.Department,
    p.WorkEmail,
    p.SipAddress,
    p.AccountName,
    p.OfficeNumber,
    p.CellPhone,
  ]
    .filter(Boolean)
    .join(" ")
    .toLowerCase();
  const words = q.split(/\s+/).filter(Boolean);
  if (mode === "exact") return hay.includes(q);
  if (mode === "all") return words.every((w) => hay.includes(w));
  return words.some((w) => hay.includes(w));
};

type SourceKey =
  | "resource"
  | "weekly"
  | "lessons"
  | "people"
  | "itblogs"
  | "marketingcal"
  | "companycal"
  | "locations"
  | "corpresp";

const ALL_SOURCES: SourceKey[] = [
  "resource",
  "weekly",
  "lessons",
  "itblogs",
  "people",
  "marketingcal",
  "companycal",
  "locations",
  "corpresp",
];

type ResultSource =
  | "ResourceLibrary"
  | "WeeklyWords"
  | "LessonsLearned"
  | "People"
  | "ITBlogs"
  | "MarketingCalendar"
  | "CompanyCalendar"
  | "Locations"
  | "CorporateResponsibilities";

interface ISearchResult {
  Title: string;
  Filename?: string;
  Path?: string;
  OriginalPath?: string;
  HitHighlightedSummary?: string;
  FileType?: string;
  ListItemId?: number;
  LastModifiedTime?: string;
  Author?: string;
  Source: ResultSource;
  PreferredName?: string;
  JobTitle?: string;
  Department?: string;
  AccountName?: string;
  SipAddress?: string;
  WorkEmail?: string;
  CellPhone?: string;
  OfficeNumber?: string;
  OriginalID?: number;
  EventDate?: string;
  EndDate?: string;
  EventLocation?: string;
  AllDayEvent?: boolean;
  Category?: string;
  RecurrenceData?: string;
  fRecurrence?: boolean;
  FreeBusy?: string;
  Overbook?: string;
  ProjectLocation?: string;
  Location?: string;
  Phone?: string;
  Fax?: string;
  Who?: string;
}

const SEARCH_SOURCES_INFO = [
  { label: "Resource Library", url: `https://vaughnconstruction.sharepoint.com/Forms/Forms/AllItems.aspx` },
  { label: "Weekly Words", url: `https://vaughnconstruction.sharepoint.com/news/Pages/default.aspx` },
  { label: "Lessons Learned", url: `https://vaughnconstruction.sharepoint.com/ll/SitePages/Default.aspx` },
  { label: "IT Blogs", url: `https://vaughnconstruction.sharepoint.com/sites/ITSite` },
  { label: "People Directory", url: `https://vaughnconstruction.sharepoint.com/SitePages/People.aspx` },
  { label: "Marketing Calendar", url: `https://vaughnconstruction.sharepoint.com/Sites/Marketing` },
  { label: "Company Calendar", url: `https://vaughnconstruction.sharepoint.com/sites/OPSSite/SitePages/Company-Calendar.aspx` },
  { label: "Locations", url: "https://vaughnconstruction.sharepoint.com/SitePages/Locations.aspx" },
  { label: "Corporate Responsibilities", url: `https://vaughnconstruction.sharepoint.com/SitePages/Corporate-Responsibilities.aspx` },
];

const sourceLabels: Record<ResultSource, string> = {
  ResourceLibrary: "Resource Library",
  WeeklyWords: "Weekly Words",
  LessonsLearned: "Lessons Learned",
  People: "People Directory",
  ITBlogs: "IT Blogs",
  MarketingCalendar: "Marketing Calendar",
  CompanyCalendar: "Company Calendar",
  Locations: "Locations",
  CorporateResponsibilities: "Corporate Responsibilities",
};

const DOWNLOADABLE_TYPES = ["docx", "xlsx", "pptx", "pdf", "msg", "eml", "doc", "xls", "ppt"];

/** ─────────────────────────────────────────────────────────────
 *  SHARE PANEL – People picker + Teams chat sender
 * ───────────────────────────────────────────────────────────── */
interface ISharePanelProps {
  targetRef: React.RefObject<HTMLElement>;
  itemTitle: string;
  itemUrl: string;
  graphClient: MSGraphClientV3 | null;
  onDismiss: () => void;
}

interface IGraphUser {
  id: string;
  displayName: string;
  mail: string;
  userPrincipalName: string;
  jobTitle?: string;
}

const SharePanel: React.FC<ISharePanelProps> = ({
  targetRef,
  itemTitle,
  itemUrl,
  graphClient,
  onDismiss,
}) => {
  const [searchText, setSearchText] = useState("");
  const [suggestions, setSuggestions] = useState<IGraphUser[]>([]);
  const [selectedPeople, setSelectedPeople] = useState<IGraphUser[]>([]);
  const [searching, setSearhing] = useState(false);
  const [sending, setSending] = useState(false);
  const [status, setStatus] = useState<{ type: "success" | "error"; msg: string } | null>(null);
  const [message, setMessage] = useState(
    `Hi, I wanted to share this with you: **${itemTitle}**\n${itemUrl}`
  );
  const debounceRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  // Search users via Graph
  const searchUsers = async (text: string) => {
    if (!graphClient || text.trim().length < 2) {
      setSuggestions([]);
      return;
    }
    setSearhing(true);
    try {
      const res: any = await graphClient
        .api(
          `/users?$filter=startswith(displayName,'${encodeURIComponent(text)}') or startswith(mail,'${encodeURIComponent(text)}')`
        )
        .select("id,displayName,mail,userPrincipalName,jobTitle")
        .top(8)
        .get();
      const users: IGraphUser[] = (res?.value || []).filter(
        (u: IGraphUser) =>
          u.mail &&
          !u.userPrincipalName?.includes("#EXT#") &&
          !selectedPeople.find((s) => s.id === u.id)
      );
      setSuggestions(users);
    } catch {
      setSuggestions([]);
    }
    setSearhing(false);
  };

  const onSearchChange = (_e: any, val?: string) => {
    const v = val || "";
    setSearchText(v);
    if (debounceRef.current) clearTimeout(debounceRef.current);
    debounceRef.current = setTimeout(() => searchUsers(v), 300);
  };

  const addPerson = (user: IGraphUser) => {
    setSelectedPeople((prev) => [...prev, user]);
    setSuggestions([]);
    setSearchText("");
  };

  const removePerson = (id: string) =>
    setSelectedPeople((prev) => prev.filter((p) => p.id !== id));

  // Send via Teams chat (1:1 per recipient)
  const sendViaTeams = async () => {
    if (!graphClient || selectedPeople.length === 0) return;
    setSending(true);
    setStatus(null);

    // ensure docs open in browser and get a sensible anchor text (file name)
    const normalizeUrlForTeams = (rawUrl: string) => {
      if (!rawUrl) return { url: rawUrl, anchor: itemTitle };
      let copyUrl = rawUrl;
      const hasWebParam = /([?&])web=1\b/.test(copyUrl);
      const isDoc = DOWNLOADABLE_TYPES.some((ext) =>
        new RegExp(`\\.(${ext})(?:$|[?#])`, "i").test(copyUrl)
      );
      if (isDoc && !hasWebParam) {
        copyUrl = copyUrl + (copyUrl.includes("?") ? "&web=1" : "?web=1");
      }
      let anchor = itemTitle;
      try {
        const p = new URL(copyUrl).pathname;
        const last = decodeURIComponent(p.substring(p.lastIndexOf("/") + 1));
        if (last) anchor = last;
      } catch {
        /* ignore - fall back to itemTitle */
      }
      return { url: copyUrl, anchor };
    };

    const normalized = normalizeUrlForTeams(itemUrl);

    const teamsMessage = {
      body: {
        contentType: "html",
        content:
          `<p>Hi! I wanted to share this with you:</p>` +
          `<p><strong>${normalized.anchor}</strong></p>` +
          `<p><a href="${normalized.url}">${normalized.anchor}</a></p>` +
          (message !== `Hi, I wanted to share this with you: **${itemTitle}**\n${itemUrl}`
            ? `<p>${message.replace(/\n/g, "<br/>")}</p>`
            : ""),
      },
    };

    const errors: string[] = [];

    try {
      // resolve current user id for proper user@odata.bind usage
      const meResp: any = await graphClient.api("/me").get();
      const meId = meResp?.id;
      if (!meId) throw new Error("Unable to resolve current user id via Graph (/me).");

      await Promise.all(
        selectedPeople.map(async (person) => {
          try {
            // Build chat payload using users('<id>') binds (Graph requires this format)
            const chatBody = {
              chatType: "oneOnOne",
              members: [
                {
                  "@odata.type": "#microsoft.graph.aadUserConversationMember",
                  // recipient should have at least one role per Graph requirements
                  roles: ["guest"],
                  "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${person.id}')`,
                },
                {
                  "@odata.type": "#microsoft.graph.aadUserConversationMember",
                  // current user becomes owner in the created chat
                  roles: ["owner"],
                  "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${meId}')`,
                },
              ],
            };

            const chat: any = await graphClient.api("/chats").post(chatBody);
            const chatId = chat?.id;
            if (!chatId) throw new Error("No chat ID returned from chat creation.");

            await graphClient.api(`/chats/${chatId}/messages`).post(teamsMessage);
          } catch (innerErr: any) {
            errors.push(`${person.displayName}: ${innerErr?.message || "failed"}`);
          }
        })
      );
    } catch (outerErr: any) {
      // top-level error (e.g. /me failed) — push for UI feedback
      errors.push(`General: ${outerErr?.message || "failed to send via Graph"}`);
    }

    setSending(false);
    if (errors.length === 0) {
      setStatus({
        type: "success",
        msg: `Sent to ${selectedPeople.map((p) => p.displayName).join(", ")} via Teams!`,
      });
      setTimeout(onDismiss, 2000);
    } else {
      setStatus({ type: "error", msg: errors.join("; ") });
    }
  };

  const getInitials = (name: string) =>
    name
      .trim()
      .split(/\s+/)
      .map((w) => w.charAt(0).toUpperCase())
      .join("")
      .substring(0, 2);

  return (
    <Callout
      target={targetRef}
      onDismiss={onDismiss}
      directionalHint={DirectionalHint.bottomAutoEdge}
      isBeakVisible={false}
      styles={{
        root: { zIndex: 9999 },
        calloutMain: {
          width: 360,
          padding: 16,
          borderRadius: 10,
          boxShadow: "0 8px 32px rgba(0,0,0,0.18)",
        },
      }}
    >
      <div onClick={(e) => e.stopPropagation()}>
        <div
          style={{
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
            marginBottom: 12,
          }}
        >
          <span style={{ fontWeight: 700, fontSize: 14 }}>Share via Teams</span>
          <IconButton
            iconProps={{ iconName: "Cancel" }}
            onClick={onDismiss}
            styles={{ root: { height: 24, width: 24 } }}
          />
        </div>

        {/* Item preview */}
        <div
          style={{
            background: "#f3f2f1",
            borderRadius: 6,
            padding: "6px 10px",
            marginBottom: 12,
            fontSize: 12,
          }}
        >
          <div
            style={{
              fontWeight: 600,
              overflow: "hidden",
              textOverflow: "ellipsis",
              whiteSpace: "nowrap",
            }}
          >
            {itemTitle}
          </div>
          <div
            style={{
              color: "#0078d4",
              overflow: "hidden",
              textOverflow: "ellipsis",
              whiteSpace: "nowrap",
            }}
          >
            {itemUrl}
          </div>
        </div>

        {/* Selected people chips */}
        {selectedPeople.length > 0 && (
          <div
            style={{
              display: "flex",
              flexWrap: "wrap",
              gap: 6,
              marginBottom: 8,
            }}
          >
            {selectedPeople.map((p) => (
              <span
                key={p.id}
                style={{
                  display: "inline-flex",
                  alignItems: "center",
                  gap: 4,
                  background: "#deecf9",
                  color: "#0078d4",
                  borderRadius: 999,
                  padding: "2px 8px 2px 6px",
                  fontSize: 12,
                  fontWeight: 600,
                }}
              >
                <span
                  style={{
                    width: 20,
                    height: 20,
                    borderRadius: "50%",
                    background: "#0078d4",
                    color: "#fff",
                    display: "inline-flex",
                    alignItems: "center",
                    justifyContent: "center",
                    fontSize: 9,
                    fontWeight: 700,
                  }}
                >
                  {getInitials(p.displayName)}
                </span>
                {p.displayName}
                <IconButton
                  iconProps={{ iconName: "Cancel" }}
                  onClick={() => removePerson(p.id)}
                  styles={{
                    root: {
                      height: 16,
                      width: 16,
                      padding: 0,
                      minWidth: 16,
                      color: "#0078d4",
                    },
                    icon: { fontSize: 10 },
                  }}
                />
              </span>
            ))}
          </div>
        )}

        {/* People search */}
        <div style={{ position: "relative", marginBottom: 8 }}>
          <TextField
            placeholder="Search people by name or email…"
            value={searchText}
            onChange={onSearchChange}
            prefix="👤"
            styles={{ root: { fontSize: 13 } }}
          />
          {searching && (
            <div style={{ padding: "4px 8px", fontSize: 12, color: "#605e5c" }}>
              Searching…
            </div>
          )}
          {suggestions.length > 0 && (
            <div
              style={{
                position: "absolute",
                top: "100%",
                left: 0,
                right: 0,
                background: "#fff",
                border: "1px solid #edebe9",
                borderRadius: 6,
                boxShadow: "0 4px 16px rgba(0,0,0,0.12)",
                zIndex: 10000,
                maxHeight: 200,
                overflowY: "auto",
              }}
            >
              {suggestions.map((u) => (
                <div
                  key={u.id}
                  onClick={() => addPerson(u)}
                  style={{
                    padding: "8px 12px",
                    cursor: "pointer",
                    display: "flex",
                    alignItems: "center",
                    gap: 8,
                    fontSize: 13,
                  }}
                  onMouseEnter={(e) =>
                    ((e.currentTarget as HTMLElement).style.background = "#f3f2f1")
                  }
                  onMouseLeave={(e) =>
                    ((e.currentTarget as HTMLElement).style.background = "transparent")
                  }
                >
                  <span
                    style={{
                      width: 28,
                      height: 28,
                      borderRadius: "50%",
                      background: "#0078d4",
                      color: "#fff",
                      display: "inline-flex",
                      alignItems: "center",
                      justifyContent: "center",
                      fontSize: 10,
                      fontWeight: 700,
                      flexShrink: 0,
                    }}
                  >
                    {getInitials(u.displayName)}
                  </span>
                  <div>
                    <div style={{ fontWeight: 600 }}>{u.displayName}</div>
                    <div style={{ fontSize: 11, color: "#605e5c" }}>
                      {u.jobTitle || u.mail}
                    </div>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>

        {/* Optional message */}
        <TextField
          label="Add a note (optional)"
          multiline
          rows={2}
          value={message}
          onChange={(_e, v) => setMessage(v || "")}
          styles={{ root: { marginBottom: 12 } }}
        />

        {status && (
          <MessageBar
            messageBarType={
              status.type === "success"
                ? MessageBarType.success
                : MessageBarType.error
            }
            styles={{ root: { marginBottom: 8, borderRadius: 6 } }}
          >
            {status.msg}
          </MessageBar>
        )}

        <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
          <DefaultButton text="Cancel" onClick={onDismiss} />
          <PrimaryButton
            text={sending ? "Sending…" : "Send via Teams"}
            disabled={selectedPeople.length === 0 || sending}
            onClick={sendViaTeams}
            iconProps={{ iconName: "TeamsLogo" }}
          />
        </div>
      </div>
    </Callout>
  );
};

/** ─────────────────────────────────────────────────────────────
 *  CARD ACTION BAR (shown on hover)
 * ───────────────────────────────────────────────────────────── */
interface ICardActionsProps {
  result: ISearchResult;
  graphClient: MSGraphClientV3 | null;
  onCardClick: () => void;
}

const CardActions: React.FC<ICardActionsProps> = ({
  result,
  graphClient,
  onCardClick,
}) => {
  const [showShare, setShowShare] = useState(false);
  const [copyFeedback, setCopyFeedback] = useState(false);
  const shareButtonRef = useRef<HTMLDivElement>(null);

  const itemUrl = (() => {
    if (result.Source === "LessonsLearned" && result.OriginalID) {
      return `${LESSONS_LEARNED_SITE_URL}/SitePages/LessonsLearnedView.aspx?LessonsLearnedId=${result.OriginalID}`;
    }
    return result.OriginalPath || result.Path || "";
  })();

  const isDownloadable =
    result.Source === "ResourceLibrary" &&
    DOWNLOADABLE_TYPES.includes((result.FileType || "").toLowerCase());

  const handleCopy = (e: React.MouseEvent<any>) => {
    e.stopPropagation();
    if (!itemUrl) return;
    // For document links ensure browser view parameter is present so link opens in browser
    let copyUrl = itemUrl;
    const hasWebParam = /([?&])web=1\b/.test(copyUrl);
    const isDocSource =
      result.Source === "ResourceLibrary" ||
      DOWNLOADABLE_TYPES.includes((result.FileType || "").toLowerCase());
    if (isDocSource && !hasWebParam) {
      copyUrl = copyUrl + (copyUrl.includes("?") ? "&web=1" : "?web=1");
    }
    navigator.clipboard.writeText(copyUrl).then(() => {
      setCopyFeedback(true);
      setTimeout(() => setCopyFeedback(false), 1800);
    });
  };

  const handleDownload = (e: React.MouseEvent<any>) => {
    e.stopPropagation();
    if (!itemUrl) return;
    const a = document.createElement("a");
    a.href = itemUrl;
    a.download = result.Filename || result.Title || "download";
    a.target = "_blank";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  // Accept an optional/loosely typed event to satisfy Fluent UI's union event types
  const handleShare = (e?: any) => {
    e?.stopPropagation();
    setShowShare((v) => !v);
  };

  if (result.Source === "People") return null;
  if (!itemUrl) return null;

  return (
    <>
      <div
        className="card-actions"
        style={{
          display: "flex",
          gap: 2,
          alignItems: "center",
        }}
        onClick={(e) => e.stopPropagation()}
      >
        {/* Copy link */}
        <TooltipHost content={copyFeedback ? "Copied!" : "Copy link"}>
          <IconButton
            iconProps={{ iconName: copyFeedback ? "CheckMark" : "Copy" }}
            ariaLabel="Copy link"
            styles={{
              root: {
                height: 28,
                width: 28,
                color: copyFeedback ? "#107c10" : "#605e5c",
              },
              icon: { fontSize: 13 },
            }}
            onClick={handleCopy}
          />
        </TooltipHost>

        {/* Download (docs only) */}
        {isDownloadable && (
          <TooltipHost content="Download">
            <IconButton
              iconProps={{ iconName: "Download" }}
              ariaLabel="Download"
              styles={{
                root: { height: 28, width: 28, color: "#605e5c" },
                icon: { fontSize: 13 },
              }}
              onClick={handleDownload}
            />
          </TooltipHost>
        )}

        {/* Share via Teams */}
        <TooltipHost content="Share via Teams">
          <div ref={shareButtonRef}>
            <IconButton
              iconProps={{ iconName: "TeamsLogo" }}
              ariaLabel="Share via Teams"
              styles={{
                root: {
                  height: 28,
                  width: 28,
                  color: showShare ? "#5558af" : "#605e5c",
                },
                icon: { fontSize: 13 },
              }}
              onClick={handleShare}
            />
          </div>
        </TooltipHost>
      </div>

      {showShare && (
        <SharePanel
          targetRef={shareButtonRef as React.RefObject<HTMLElement>}
          itemTitle={result.Title || result.Filename || "(no title)"}
          itemUrl={itemUrl}
          graphClient={graphClient}
          onDismiss={() => setShowShare(false)}
        />
      )}
    </>
  );
};

/** ─────────────────────────────────────────────────────────────
 *  Photo / Persona cache
 * ───────────────────────────────────────────────────────────── */
const peoplePhotoCache = new Map<
  string,
  { url: string | null; loading: boolean; error: boolean }
>();

const PeoplePersonaWithLive: React.FC<{
  context: any;
  graphClient: MSGraphClientV3 | null;
  displayName: string;
  upnOrEmail: string;
  jobTitle?: string;
  department?: string;
}> = ({ context, graphClient, displayName, upnOrEmail, jobTitle, department }) => {
  const normalized = normalizeUpn(upnOrEmail);

  const [photoState, setPhotoState] = useState<{
    url: string | null;
    loading: boolean;
    error: boolean;
  }>({ url: null, loading: false, error: false });

  const mountedRef = useRef(true);
  const loadingRef = useRef(false);

  const getInitials = (name: string): string => {
    if (!name) return "??";
    return name
      .trim()
      .split(/\s+/)
      .map((w) => w?.charAt(0).toUpperCase())
      .join("")
      .substring(0, 2);
  };

  useEffect(() => {
    if (!normalized) return;
    const cached = peoplePhotoCache.get(normalized);
    if (cached) setPhotoState(cached);
  }, [normalized]);

  useEffect(() => {
    const loadPhoto = async () => {
      if (!graphClient || !normalized || loadingRef.current) return;
      const cached = peoplePhotoCache.get(normalized);
      if (cached && (cached.url || cached.error)) {
        setPhotoState(cached);
        return;
      }
      if (cached?.loading) return;
      loadingRef.current = true;
      const loadingState = { url: null, loading: true, error: false };
      peoplePhotoCache.set(normalized, loadingState);
      setPhotoState(loadingState);
      try {
        const blob = await graphClient
          .api(`/users/${normalized}/photo/$value`)
          .responseType(ResponseType.BLOB)
          .get();
        if (!mountedRef.current) return;
        const url = URL.createObjectURL(blob as Blob);
        const ok = { url, loading: false, error: false };
        peoplePhotoCache.set(normalized, ok);
        setPhotoState(ok);
      } catch {
        const fail = { url: null, loading: false, error: true };
        peoplePhotoCache.set(normalized, fail);
        if (mountedRef.current) setPhotoState(fail);
      } finally {
        loadingRef.current = false;
      }
    };
    loadPhoto();
  }, [graphClient, normalized]);

  useEffect(() => {
    return () => { mountedRef.current = false; };
  }, []);

  const personaElement = (
    <Persona
      text={displayName}
      secondaryText={jobTitle}
      tertiaryText={department}
      size={PersonaSize.size40}
      imageUrl={photoState.url || undefined}
      imageInitials={photoState.url ? undefined : getInitials(displayName)}
      coinSize={40}
    />
  );

  if (context?.serviceScope && normalized && !photoState.error) {
    return (
      <div style={{ position: "relative" }}>
        {personaElement}
        <div style={{ position: "absolute", inset: 0, opacity: 0, pointerEvents: "all" }}>
          <LivePersona serviceScope={context.serviceScope} upn={normalized} template={personaElement} />
        </div>
      </div>
    );
  }

  return personaElement;
};

/** ─────────────────────────────────────────────────────────────
 *  HTML sanitize helpers
 * ───────────────────────────────────────────────────────────── */
const normalizeSearchSummaryHtml = (html: string) => {
  if (!html) return "";
  return html
    .replace(/<ddd\s*\/>/gi, "… ")
    .replace(/<c0>/gi, '<mark style="background:yellow">')
    .replace(/<\/c0>/gi, "</mark>")
    .replace(/<c\d+>/gi, '<mark style="background:yellow">')
    .replace(/<\/c\d+>/gi, "</mark>");
};

const sanitizeHtml = (html: string) =>
  DOMPurify.sanitize(html, {
    USE_PROFILES: { html: true },
    ADD_TAGS: ["mark"],
    ADD_ATTR: ["style"],
  });

const getSafeHtmlForTooltip = (r: ISearchResult, words: string[]) => {
  if (r.Source === "WeeklyWords" || r.Source === "LessonsLearned") {
    return sanitizeHtml(r.HitHighlightedSummary || "");
  }
  if (r.Source === "ITBlogs") {
    return sanitizeHtml(normalizeSearchSummaryHtml(r.HitHighlightedSummary || ""));
  }
  return sanitizeHtml(highlight(r.HitHighlightedSummary || "", words));
};

/** ─────────────────────────────────────────────────────────────
 *  MAIN COMPONENT
 * ───────────────────────────────────────────────────────────── */
const SearchDocuments: React.FC<ISearchDocumentsProps> = ({
  context,
  msGraphClientFactory,
  variant = "SearchWithResult",
}) => {
  const spRoot = useMemo(() => spfi(SITE_URL).using(SPFx(context)), [context]);

  const [graphClient, setGraphClient] = useState<MSGraphClientV3 | null>(null);

  useEffect(() => {
    const initGraph = async () => {
      if (!msGraphClientFactory) return;
      const client = await msGraphClientFactory.getClient("3");
      setGraphClient(client);
    };
    initGraph();
  }, [msGraphClientFactory]);

  const [query, setQuery] = useState("");
  const [mode, setMode] = useState<Mode>("exact");
  const [results, setResults] = useState<ISearchResult[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string>("");

  const [fileTypeFilters, setFileTypeFilters] = useState<string[]>([]);
  const [availableFileTypes, setAvailableFileTypes] = useState<string[]>([
    "docx",
    "xlsx",
    "pptx",
    "pdf",
    "msg",
    "eml",
  ]);

  const [sourceFilters, setSourceFilters] = useState<SourceKey[]>(ALL_SOURCES);

  const pageSize = 10;
  const [page, setPage] = useState(1);
  const totalPages = Math.ceil(results.length / pageSize);
  const paginatedResults = results.slice((page - 1) * pageSize, page * pageSize);

  const words = query.trim() ? query.trim().split(/\s+/) : [];

  const modeOptions: IChoiceGroupOption[] = [
    { key: "exact", text: "Exact phrase" },
    { key: "all", text: "All words" },
    { key: "any", text: "Any word" },
  ];

  const toggleFileType = (ext: string, checked?: boolean) => {
    setPage(1);
    setFileTypeFilters((prev) => {
      const exists = prev.includes(ext);
      if (checked && !exists) return [...prev, ext];
      if (!checked && exists) return prev.filter((e) => e !== ext);
      return prev;
    });
  };

  const toggleSource = (key: SourceKey, checked?: boolean) => {
    setPage(1);
    setSourceFilters((prev) => {
      const exists = prev.includes(key);
      if (checked && !exists) return [...prev, key];
      if (!checked && exists) return prev.filter((k) => k !== key);
      return prev;
    });
  };

  const handleAllSourcesToggle = (_e: any, checked?: boolean) => {
    setPage(1);
    setSourceFilters(checked ? ALL_SOURCES : []);
  };

  const sourceCounts = useMemo(() => {
    const map: Record<string, number> = {};
    results.forEach((r) => {
      map[r.Source] = (map[r.Source] || 0) + 1;
    });
    return map;
  }, [results]);

  const siteIdCache = useRef(new Map<string, string>());

  const getGraphSiteId = async (absoluteSiteUrl: string): Promise<string> => {
    if (!graphClient) throw new Error("Graph client not initialized.");
    const u = new URL(absoluteSiteUrl);
    const hostname = u.hostname;
    const path = u.pathname.replace(/\/$/, "");
    const key = `${hostname}:${path || "/"}`;
    const cached = siteIdCache.current.get(key);
    if (cached) return cached;
    const resp: any = await graphClient
      .api(`/sites/${hostname}:${path || "/"}`)
      .get();
    const id = resp?.id;
    if (!id) throw new Error(`Unable to resolve site id for ${absoluteSiteUrl}`);
    siteIdCache.current.set(key, id);
    return id;
  };

  const graphSearchListItems = async (
    queryString: string,
    listId: string
  ): Promise<any[]> => {
    if (!graphClient) return [];
    const body = {
      requests: [
        {
          entityTypes: ["listItem"],
          query: { queryString },
          from: 0,
          size: 50,
          fields: [
            "Title",
            "webUrl",
            "lastModifiedDateTime",
            "createdBy",
            "sharepointIds",
          ],
        },
      ],
    };
    const res: any = await graphClient.api("/search/query").post(body);
    const hits = res?.value?.[0]?.hitsContainers?.[0]?.hits ?? [];
    const target = normalizeGuid(listId);
    return hits
      .map((h: any) => h.resource)
      .filter((r: any) => normalizeGuid(r?.sharepointIds?.listId) === target);
  };

  const graphGetListItemFields = async (
    siteUrl: string,
    listId: string,
    listItemId: number,
    fieldsSelect: string[]
  ) => {
    if (!graphClient) throw new Error("Graph client not initialized.");
    const siteId = await getGraphSiteId(siteUrl);
    const selectPart = fieldsSelect.join(",");
    const item: any = await graphClient
      .api(
        `/sites/${siteId}/lists/${listId}/items/${listItemId}?expand=fields($select=${selectPart})`
      )
      .get();
    const fieldsObj = item?.fields || {};
    const createdBy =
      item?.createdByUser?.displayName ||
      item?.createdBy?.user?.displayName ||
      "";
    const lastModifiedBy =
      item?.lastModifiedByUser?.displayName ||
      item?.lastModifiedBy?.user?.displayName ||
      "";
    return { fields: fieldsObj, createdBy, lastModifiedBy };
  };

  const runSearch = async (overrideText?: string, overrideMode?: Mode) => {
    const textToUse = overrideText !== undefined ? overrideText : query;
    const modeToUse = overrideMode !== undefined ? overrideMode : mode;
    const trimmed = textToUse.trim();
    if (!trimmed) {
      setResults([]);
      setError("");
      return;
    }

    if (variant === "SearchBar") {
      const url = `${RESULTS_PAGE_URL}?q=${encodeURIComponent(trimmed)}`;
      window.open(url, "_blank");
      return;
    }

    setLoading(true);
    setError("");
    setPage(1);

    try {
      const kqlBody = buildKql(trimmed, modeToUse, fields);
      const graphQueryString = buildGraphQueryString(trimmed, modeToUse);
      const activeSources: SourceKey[] =
        sourceFilters.length > 0 ? sourceFilters : ALL_SOURCES;
      const needGraph = activeSources.some((s) =>
        [
          "weekly",
          "lessons",
          "itblogs",
          "marketingcal",
          "companycal",
          "locations",
          "corpresp",
        ].includes(s)
      );

      if (needGraph && !graphClient) {
        setLoading(false);
        setError("Graph client is not initialized yet. Please try again.");
        return;
      }

      const searchPromises: Promise<ISearchResult[]>[] = [];

      // Resource Library
      if (activeSources.includes("resource")) {
        const docsPromise = (async (): Promise<ISearchResult[]> => {
          let filter = `(${kqlBody}) AND (contentclass:STS_ListItem_DocumentLibrary AND IsDocument:"true" AND ListId:"${DOCUMENT_LIBRARY_ID}")`;
          if (fileTypeFilters.length > 0) {
            const ftClause = fileTypeFilters
              .map((ft) => `FileType:${ft}`)
              .join(" OR ");
            filter += ` AND (${ftClause})`;
          }
          const res: any = await spRoot.search({
            Querytext: filter,
            RowLimit: 500,
            TrimDuplicates: false,
            SelectProperties: [
              "Title",
              "Path",
              "OriginalPath",
              "FileType",
              "Filename",
              "ListItemId",
              "LastModifiedTime",
              "Author",
              "HitHighlightedSummary",
            ],
            EnableQueryRules: true,
            SortList: [{ Property: "Rank", Direction: 1 }],
          });
          const items: ISearchResult[] = (
            res.PrimarySearchResults || []
          ).map((i: any) => ({
            Title: i.Title || "",
            Filename: i.Filename || "",
            Path: i.Path,
            OriginalPath: i.OriginalPath,
            HitHighlightedSummary: i.HitHighlightedSummary,
            FileType: (i.FileType || "").toLowerCase(),
            ListItemId: Number(i.ListItemId),
            LastModifiedTime: i.LastModifiedTime,
            Author: i.Author,
            Source: "ResourceLibrary",
          }));
          const dynamicTypes: string[] = Array.from(
            new Set(
              items
                .map((it) => it.FileType)
                .filter((x): x is string => !!x && x.trim().length > 0)
            )
          );
          if (dynamicTypes.length > 0) setAvailableFileTypes(dynamicTypes);
          return items;
        })();
        searchPromises.push(docsPromise);
      }

      // IT Blogs
      if (activeSources.includes("itblogs") && graphClient) {
        const itBlogsPromise = (async (): Promise<ISearchResult[]> => {
          const body = {
            requests: [
              {
                entityTypes: ["listItem"],
                query: { queryString: graphQueryString },
                from: 0,
                size: 50,
                fields: [
                  "Title",
                  "webUrl",
                  "lastModifiedDateTime",
                  "createdBy",
                  "sharepointIds",
                ],
              },
            ],
          };
          const res: any = await graphClient.api("/search/query").post(body);
          const hits = res?.value?.[0]?.hitsContainers?.[0]?.hits ?? [];
          const targetListId = normalizeGuid(IT_BLOGS_LIBRARY_ID);
          const filtered: any[] = hits
            .map((h: any) => h.resource)
            .filter((r: any) => {
              const listId = normalizeGuid(r?.sharepointIds?.listId);
              if (listId && listId === targetListId) return true;
              return (r?.webUrl || "")
                .toLowerCase()
                .includes(IT_BLOGS_PATH_CONTAINS.toLowerCase());
            });
          if (filtered.length === 0) return [];
          const spIT = spfi(IT_SITE_URL).using(SPFx(context));
          const itemsWithDetails = await Promise.all(
            filtered.map(async (r: any) => {
              const id = parseInt(r?.sharepointIds?.listItemId || "0", 10);
              if (!id) return null;
              try {
                const item: any = await spIT.web.lists
                  .getById(IT_BLOGS_LIBRARY_ID)
                  .items.getById(id)
                  .select("Id", "Title", "CanvasContent1", "Modified", "Author/Title")
                  .expand("Author")();
                return {
                  Title: item.Title || r?.fields?.Title || "(no title)",
                  Filename: "",
                  Path: r.webUrl,
                  OriginalPath: "",
                  HitHighlightedSummary: item.CanvasContent1 || "",
                  FileType: "aspx",
                  ListItemId: item.Id,
                  LastModifiedTime: r.lastModifiedDateTime || item.Modified || "",
                  Author: item.Author?.Title || r.createdBy?.user?.displayName || "",
                  Source: "ITBlogs" as const,
                } as ISearchResult;
              } catch {
                return {
                  Title: r?.fields?.Title || "(no title)",
                  Filename: "",
                  Path: r.webUrl,
                  OriginalPath: "",
                  HitHighlightedSummary: "",
                  FileType: "aspx",
                  ListItemId: 0,
                  LastModifiedTime: r.lastModifiedDateTime || "",
                  Author: r.createdBy?.user?.displayName || "",
                  Source: "ITBlogs" as const,
                } as ISearchResult;
              }
            })
          );
          return itemsWithDetails.filter((x): x is ISearchResult => x !== null);
        })();
        searchPromises.push(itBlogsPromise);
      }

      // Weekly Words
      if (activeSources.includes("weekly") && graphClient) {
        const weeklyPromise = (async (): Promise<ISearchResult[]> => {
          const body = {
            requests: [
              {
                entityTypes: ["listItem"],
                query: { queryString: graphQueryString },
                from: 0,
                size: 50,
                fields: [
                  "Title",
                  "ArticleDate",
                  "VaughnContent",
                  "Abstract",
                  "PublishingContact",
                  "webUrl",
                  "lastModifiedDateTime",
                  "createdBy",
                  "sharepointIds",
                  "OriginalID",
                ],
              },
            ],
          };
          const res: any = await graphClient.api("/search/query").post(body);
          const hits = res?.value?.[0]?.hitsContainers?.[0]?.hits ?? [];
          const filtered: any[] = hits
            .map((h: any) => h.resource)
            .filter(
              (r: any) =>
                normalizeGuid(r?.sharepointIds?.listId) ===
                normalizeGuid(WEEKLY_WORDS_LIST_ID)
            );
          if (filtered.length === 0) return [];
          const spWeekly = spfi(WEEKLY_WORDS_SITE_URL).using(SPFx(context));
          const itemsWithDetails = await Promise.all(
            filtered.map(async (r: any) => {
              const id = parseInt(r.sharepointIds?.listItemId || "0", 10);
              if (!id) return null;
              try {
                const item: any = await spWeekly.web.lists
                  .getById(WEEKLY_WORDS_LIST_ID)
                  .items.getById(id)
                  .select(
                    "Id",
                    "Title",
                    "ArticleDate",
                    "VaughnContent",
                    "Abstract",
                    "PublishingContact/Title",
                    "PublishingContact/Id",
                    "OriginalID",
                    "Modified"
                  )
                  .expand("PublishingContact")();
                return {
                  Title: item.Title || r.fields?.Title || "(no title)",
                  Filename: "",
                  Path: r.webUrl,
                  OriginalPath: "",
                  HitHighlightedSummary: item.VaughnContent || item.Abstract || "",
                  FileType: "item",
                  ListItemId: item.Id,
                  OriginalID: item.OriginalID ? Number(item.OriginalID) : undefined,
                  LastModifiedTime: item.Modified || r.lastModifiedDateTime || "",
                  Author:
                    item.PublishingContact?.Title ||
                    r.createdBy?.user?.displayName ||
                    "",
                  Source: "WeeklyWords" as const,
                } as ISearchResult;
              } catch {
                return null;
              }
            })
          );
          return itemsWithDetails.filter((x): x is ISearchResult => x !== null);
        })();
        searchPromises.push(weeklyPromise);
      }

      // Lessons Learned
      if (activeSources.includes("lessons") && graphClient) {
        const lessonsPromise = (async (): Promise<ISearchResult[]> => {
          const body = {
            requests: [
              {
                entityTypes: ["listItem"],
                query: { queryString: graphQueryString },
                from: 0,
                size: 50,
                fields: [
                  "Title",
                  "Body",
                  "Contact",
                  "webUrl",
                  "lastModifiedDateTime",
                  "createdBy",
                  "sharepointIds",
                  "OriginalID",
                ],
              },
            ],
          };
          const res: any = await graphClient.api("/search/query").post(body);
          const hits = res?.value?.[0]?.hitsContainers?.[0]?.hits ?? [];
          const targetListId = normalizeGuid(LESSONS_LEARNED_LIST_ID);
          const filtered: any[] = hits
            .map((h: any) => h.resource)
            .filter((r: any) => {
              const listId = normalizeGuid(r?.sharepointIds?.listId);
              if (listId && listId === targetListId) return true;
              return (r?.webUrl || "").toLowerCase().includes("/ll/");
            });
          if (filtered.length === 0) return [];
          const spLessons = spfi(LESSONS_LEARNED_SITE_URL).using(SPFx(context));
          const itemsWithDetails = await Promise.all(
            filtered.map(async (r: any) => {
              const id = parseInt(r.sharepointIds?.listItemId || "0", 10);
              if (!id) return null;
              try {
                const item: any = await spLessons.web.lists
                  .getById(LESSONS_LEARNED_LIST_ID)
                  .items.getById(id)
                  .select(
                    "Id",
                    "Title",
                    "PublishedDate",
                    "Body",
                    "Contact/Title",
                    "OriginalID",
                    "Modified"
                  )
                  .expand("Contact")();
                return {
                  Title: item.Title || r.fields?.Title || "(no title)",
                  Filename: "",
                  Path: r.webUrl,
                  OriginalPath: "",
                  HitHighlightedSummary: item.Body || "",
                  FileType: "item",
                  ListItemId: item.Id,
                  OriginalID: item.OriginalID ? Number(item.OriginalID) : undefined,
                  LastModifiedTime:
                    item.Modified || item.PublishedDate || r.lastModifiedDateTime || "",
                  Author: item.Contact?.Title || r.createdBy?.user?.displayName || "",
                  Source: "LessonsLearned" as const,
                } as ISearchResult;
              } catch {
                return null;
              }
            })
          );
          return itemsWithDetails.filter((x): x is ISearchResult => x !== null);
        })();
        searchPromises.push(lessonsPromise);
      }

      // People
      if (activeSources.includes("people")) {
        const peoplePromise = (async (): Promise<ISearchResult[]> => {
          const peopleQuery = {
            Querytext: trimmed,
            SourceId: PEOPLE_SOURCE_ID,
            RowLimit: 500,
            RowsPerPage: 500,
            TrimDuplicates: true,
            SelectProperties: [
              "SipAddress",
              "PreferredName",
              "JobTitle",
              "Department",
              "WorkEmail",
              "CellPhone",
              "OfficeNumber",
              "AccountName",
            ],
          };
          const res: any = await spRoot.search(peopleQuery);
          const items: ISearchResult[] = (res.PrimarySearchResults || []).map(
            (i: any) => ({
              Title: i.PreferredName || "(no name)",
              Filename: "",
              Path: i.WorkEmail || i.SipAddress || "",
              OriginalPath: "",
              FileType: "person",
              ListItemId: 0,
              LastModifiedTime: "",
              Author: "",
              Source: "People",
              PreferredName: i.PreferredName,
              JobTitle: i.JobTitle,
              Department: i.Department,
              AccountName: i.AccountName,
              SipAddress: i.SipAddress,
              WorkEmail: i.WorkEmail,
              CellPhone: i.CellPhone,
              OfficeNumber: i.OfficeNumber,
            })
          );
          return items
            .filter((p) => !isGuestPerson(p))
            .filter((p) => peopleMatchesQuery(p, trimmed, modeToUse));
        })();
        searchPromises.push(peoplePromise);
      }

      const calendarFieldSelect = [
        "Id",
        "Title",
        "Location",
        "EventDate",
        "RecurrenceData",
        "fRecurrence",
        "fAllDayEvent",
        "EndDate",
        "Description",
        "Category",
        "FreeBusy",
        "Overbook",
        "Modified",
        "Created",
        "ParticipantsPicker",
      ];

      // Company Calendar
      if (activeSources.includes("companycal") && graphClient) {
        const p = (async (): Promise<ISearchResult[]> => {
          const hits = await graphSearchListItems(
            graphQueryString,
            COMPANY_CAL_LIST_ID
          );
          if (!hits.length) return [];
          const items = await Promise.all(
            hits.map(async (r: any) => {
              const id = parseInt(r?.sharepointIds?.listItemId || "0", 10);
              if (!id) return null;
              try {
                const { fields: f, createdBy, lastModifiedBy } =
                  await graphGetListItemFields(
                    COMPANY_CAL_SITE_URL,
                    COMPANY_CAL_LIST_ID,
                    id,
                    calendarFieldSelect
                  );
                return {
                  Title: f?.Title || r?.fields?.Title || "(no title)",
                  Path: r?.webUrl,
                  FileType: "event",
                  ListItemId: id,
                  LastModifiedTime: f?.Modified || r?.lastModifiedDateTime || "",
                  Author:
                    lastModifiedBy ||
                    createdBy ||
                    r?.createdBy?.user?.displayName ||
                    "",
                  Source: "CompanyCalendar" as const,
                  EventDate: f?.EventDate,
                  EndDate: f?.EndDate,
                  EventLocation: f?.Location,
                  AllDayEvent: f?.fAllDayEvent,
                } as ISearchResult;
              } catch {
                return {
                  Title: r?.fields?.Title || "(no title)",
                  Path: r?.webUrl,
                  FileType: "event",
                  ListItemId: id,
                  LastModifiedTime: r?.lastModifiedDateTime || "",
                  Author: r?.createdBy?.user?.displayName || "",
                  Source: "CompanyCalendar" as const,
                } as ISearchResult;
              }
            })
          );
          return items.filter(Boolean) as ISearchResult[];
        })();
        searchPromises.push(p);
      }

      // Marketing Calendar
      if (activeSources.includes("marketingcal") && graphClient) {
        const p = (async (): Promise<ISearchResult[]> => {
          const hits = await graphSearchListItems(
            graphQueryString,
            MARKETING_CAL_LIST_ID
          );
          if (!hits.length) return [];
          const items = await Promise.all(
            hits.map(async (r: any) => {
              const id = parseInt(r?.sharepointIds?.listItemId || "0", 10);
              if (!id) return null;
              try {
                const { fields: f, createdBy, lastModifiedBy } =
                  await graphGetListItemFields(
                    MARKETING_CAL_SITE_URL,
                    MARKETING_CAL_LIST_ID,
                    id,
                    calendarFieldSelect
                  );
                return {
                  Title: f?.Title || r?.fields?.Title || "(no title)",
                  Path: r?.webUrl,
                  FileType: "event",
                  ListItemId: id,
                  LastModifiedTime: f?.Modified || r?.lastModifiedDateTime || "",
                  Author:
                    lastModifiedBy ||
                    createdBy ||
                    r?.createdBy?.user?.displayName ||
                    "",
                  Source: "MarketingCalendar" as const,
                  EventDate: f?.EventDate,
                  EndDate: f?.EndDate,
                  EventLocation: f?.Location,
                  AllDayEvent: f?.fAllDayEvent,
                  Category: f?.Category,
                  RecurrenceData: f?.RecurrenceData,
                  fRecurrence: f?.fRecurrence,
                  FreeBusy: f?.FreeBusy,
                  Overbook: f?.Overbook,
                } as ISearchResult;
              } catch {
                return {
                  Title: r?.fields?.Title || "(no title)",
                  Path: r?.webUrl,
                  FileType: "event",
                  ListItemId: id,
                  LastModifiedTime: r?.lastModifiedDateTime || "",
                  Author: r?.createdBy?.user?.displayName || "",
                  Source: "MarketingCalendar" as const,
                } as ISearchResult;
              }
            })
          );
          return items.filter((x): x is ISearchResult => x !== null);
        })();
        searchPromises.push(p);
      }

      // Locations
      if (activeSources.includes("locations") && graphClient) {
        const p = (async (): Promise<ISearchResult[]> => {
          const hits = await graphSearchListItems(
            graphQueryString,
            LOCATIONS_LIST_ID
          );
          if (!hits.length) return [];
          const fieldsSelect = [
            "Id",
            "Title",
            "Project_x0020_Location",
            "Phone",
            "Fax",
            "Location",
          ];
          const items = await Promise.all(
            hits.map(async (r: any) => {
              const id = parseInt(r?.sharepointIds?.listItemId || "0", 10);
              if (!id) return null;
              try {
                const { fields: f, createdBy, lastModifiedBy } =
                  await graphGetListItemFields(
                    LOCATIONS_SITE_URL,
                    LOCATIONS_LIST_ID,
                    id,
                    fieldsSelect
                  );
                return {
                  Title: f?.Title || r?.fields?.Title || "(no title)",
                  Path: r?.webUrl,
                  FileType: "item",
                  ListItemId: id,
                  LastModifiedTime: "",
                  Author: lastModifiedBy || createdBy || "",
                  Source: "Locations" as const,
                  ProjectLocation: f?.Project_x0020_Location,
                  Phone: f?.Phone,
                  Fax: f?.Fax,
                  Location: f?.Location,
                } as ISearchResult;
              } catch {
                return {
                  Title: r?.fields?.Title || "(no title)",
                  Path: r?.webUrl,
                  FileType: "item",
                  ListItemId: id,
                  Source: "Locations" as const,
                } as ISearchResult;
              }
            })
          );
          return items.filter((x): x is ISearchResult => x !== null);
        })();
        searchPromises.push(p);
      }

      // Corporate Responsibilities
      if (activeSources.includes("corpresp") && graphClient) {
        const p = (async (): Promise<ISearchResult[]> => {
          const hits = await graphSearchListItems(
            graphQueryString,
            CORP_RESP_LIST_ID
          );
          if (!hits.length) return [];
          const fieldsSelect = ["Id", "Title", "Who"];
          const items = await Promise.all(
            hits.map(async (r: any) => {
              const id = parseInt(r?.sharepointIds?.listItemId || "0", 10);
              if (!id) return null;
              try {
                const { fields: f, createdBy, lastModifiedBy } =
                  await graphGetListItemFields(
                    CORP_RESP_SITE_URL,
                    CORP_RESP_LIST_ID,
                    id,
                    fieldsSelect
                  );
                return {
                  Title: f?.Title || r?.fields?.Title || "(no title)",
                  Path: r?.webUrl,
                  FileType: "item",
                  ListItemId: id,
                  Source: "CorporateResponsibilities" as const,
                  Who: f?.Who,
                  Author: lastModifiedBy || createdBy || "",
                } as ISearchResult;
              } catch {
                return {
                  Title: r?.fields?.Title || "(no title)",
                  Path: r?.webUrl,
                  FileType: "item",
                  ListItemId: id,
                  Source: "CorporateResponsibilities" as const,
                } as ISearchResult;
              }
            })
          );
          return items.filter((x): x is ISearchResult => x !== null);
        })();
        searchPromises.push(p);
      }

      const resultArrays = await Promise.all(searchPromises);
      const merged: ISearchResult[] = ([] as ISearchResult[]).concat(
        ...resultArrays
      );

      const lowerQuery = trimmed.toLowerCase();
      const queryWords = lowerQuery.split(/\s+/).filter(Boolean);

      const scored = merged
        .map((it, index) => {
          const title = (it.Title || "").toLowerCase();
          const secondary = (
            it.Filename ||
            it.Department ||
            it.JobTitle ||
            it.EventLocation ||
            it.Category ||
            it.ProjectLocation ||
            it.Location ||
            it.Who ||
            ""
          ).toLowerCase();
          let score = 0;
          if (modeToUse === "exact") {
            if (title.includes(lowerQuery)) score += 100;
            if (secondary.includes(lowerQuery)) score += 50;
          } else {
            let t = 0;
            let s = 0;
            queryWords.forEach((w) => {
              if (w && title.includes(w)) t++;
              if (w && secondary.includes(w)) s++;
            });
            if (modeToUse === "all") {
              score += t * 10 + s * 5;
              if (t === queryWords.length && queryWords.length > 0) score += 40;
            } else {
              score += t * 5 + s * 2;
            }
          }
          return { it, score, index };
        })
        .sort((a, b) =>
          b.score !== a.score ? b.score - a.score : a.index - b.index
        )
        .map((x) => x.it);

      setResults(scored);
    } catch (e: any) {
      console.error(e);
      setError(e?.message || "Search failed.");
    }

    setLoading(false);
  };

  const autoRunRef = useRef<{ q: string; ran: boolean }>({ q: "", ran: false });

  useEffect(() => {
    if (variant !== "SearchWithResult") return;
    let qParam = "";
    try {
      const params = new URLSearchParams(window.location.search);
      qParam = params.get("q") || "";
    } catch {
      qParam = "";
    }
    const trimmed = qParam.trim();
    if (!trimmed) return;
    const activeSources: SourceKey[] =
      sourceFilters.length > 0 ? sourceFilters : ALL_SOURCES;
    const needGraph = activeSources.some((s) =>
      [
        "weekly",
        "lessons",
        "itblogs",
        "marketingcal",
        "companycal",
        "locations",
        "corpresp",
      ].includes(s)
    );
    if (needGraph && !graphClient) return;
    if (autoRunRef.current.ran && autoRunRef.current.q === trimmed) return;
    autoRunRef.current = { q: trimmed, ran: true };
    setQuery(trimmed);
    setMode("exact");
    runSearch(trimmed, "exact");
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [variant, graphClient]);

  const openResult = (item: ISearchResult) => {
    if (item.Source === "People") {
      if (item.WorkEmail) window.location.href = `mailto:${item.WorkEmail}`;
      else if (item.SipAddress) window.open(`sip:${item.SipAddress}`, "_blank");
      return;
    }
    if (item.Source === "LessonsLearned" && item.OriginalID) {
      const url = `${LESSONS_LEARNED_SITE_URL}/SitePages/LessonsLearnedView.aspx?LessonsLearnedId=${item.OriginalID}`;
      window.open(url, "_blank");
      return;
    }
    const basePath = item.OriginalPath || item.Path;
    if (!basePath) return;
    const isPage =
      (item.FileType || "").toLowerCase() === "aspx" ||
      basePath.toLowerCase().includes(".aspx");
    const url = isPage ? basePath : `${basePath}?web=1`;
    window.open(url, "_blank");
  };

  /** Styles */
  const cardStyle: React.CSSProperties = {
    background: "#ffffff",
    borderRadius: 10,
    boxShadow: "0 1px 3px rgba(0,0,0,0.08)",
    padding: 12,
    position: "relative",
  };

  const pillStyle: React.CSSProperties = {
    display: "inline-block",
    padding: "2px 8px",
    borderRadius: 999,
    fontSize: 10,
    fontWeight: 600,
    textTransform: "uppercase",
    background: "#f3f2f1",
    color: "#605e5c",
    lineHeight: 1.6,
  };

  const metaRowStyle: React.CSSProperties = {
    fontSize: 12,
    color: "#605e5c",
    marginTop: 6,
    display: "flex",
    flexWrap: "wrap",
    gap: 8,
  };

  // SearchBar variant
  if (variant === "SearchBar") {
    return (
      <div style={{ padding: 20 }}>
        <Stack tokens={{ childrenGap: 8 }} styles={{ root: { maxWidth: 600 } }}>
          <div style={{ display: "flex", alignItems: "flex-end", gap: 6 }}>
            <TextField
              label="Search"
              placeholder="Type search text…"
              value={query}
              onChange={(_e, v) => setQuery(v || "")}
              onKeyDown={(e) => {
                if (e.key === "Enter") {
                  e.preventDefault();
                  runSearch();
                }
              }}
              styles={{ root: { flex: 1 } }}
            />
            <TooltipHost
              content={
                <div style={{ maxWidth: 260 }}>
                  <div style={{ fontWeight: 600, marginBottom: 6 }}>
                    Searches across:
                  </div>
                  <ul style={{ paddingLeft: 16, margin: 0 }}>
                    {SEARCH_SOURCES_INFO.map((s) => (
                      <li key={s.label}>{s.label}</li>
                    ))}
                  </ul>
                </div>
              }
            >
              <IconButton
                iconProps={{ iconName: "Info" }}
                ariaLabel="Search sources info"
                styles={{ root: { marginBottom: 2 } }}
              />
            </TooltipHost>
          </div>
          <PrimaryButton text="Search" onClick={() => runSearch()} />
        </Stack>
      </div>
    );
  }

  const allSourcesSelected = sourceFilters.length === ALL_SOURCES.length;

  return (
    <>
      {/* Scoped hover CSS */}
      <style>{`
        .vc-result-card .card-actions { opacity: 0; transition: opacity 0.15s ease; }
        .vc-result-card:hover .card-actions { opacity: 1; }
        .vc-result-card:hover { box-shadow: 0 4px 14px rgba(0,0,0,0.13) !important; }
      `}</style>

      <div style={{ padding: 20, background: "#faf9f8" }}>
        <div style={{ maxWidth: 1200, margin: "0 auto" }}>
          {/* Header */}
          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              alignItems: "flex-start",
              marginBottom: 16,
            }}
          >
            <div>
              <h2 style={{ margin: 0 }}>Advanced Search</h2>
              <p style={{ margin: "4px 0 6px", color: "#605e5c" }}>
                This search looks across the following sources:
              </p>
              <div
                style={{
                  fontSize: 12,
                  display: "flex",
                  flexWrap: "wrap",
                  gap: 8,
                }}
              >
                {SEARCH_SOURCES_INFO.map((s) => (
                  <a
                    key={s.label}
                    href={s.url}
                    target="_blank"
                    rel="noopener noreferrer"
                    style={{
                      color: "#0078d4",
                      textDecoration: "none",
                      fontWeight: 500,
                    }}
                  >
                    {s.label}
                  </a>
                ))}
              </div>
            </div>

            {results.length > 0 && (
              <div style={cardStyle}>
                <div
                  style={{ fontSize: 12, fontWeight: 600, marginBottom: 6 }}
                >
                  Result summary
                </div>
                <div style={{ fontSize: 12 }}>
                  <div style={{ marginBottom: 4 }}>
                    Total:{" "}
                    <strong>
                      {results.length} item{results.length !== 1 ? "s" : ""}
                    </strong>
                  </div>
                  <div
                    style={{ display: "flex", flexWrap: "wrap", gap: 6 }}
                  >
                    {Object.entries(sourceCounts).map(([key, count]) => (
                      <span key={key} style={pillStyle}>
                        {sourceLabels[key as ResultSource]}: {count}
                      </span>
                    ))}
                  </div>
                </div>
              </div>
            )}
          </div>

          {/* Search controls */}
          <div
            style={{
              display: "grid",
              gridTemplateColumns: "minmax(0, 2fr) minmax(0, 1.4fr)",
              gap: 16,
              alignItems: "flex-start",
            }}
          >
            <div style={cardStyle}>
              <Stack tokens={{ childrenGap: 8 }}>
                <TextField
                  label="Search text"
                  placeholder="Type keywords…"
                  value={query}
                  onChange={(_e, v) => setQuery(v || "")}
                  onKeyDown={(e) => {
                    if (e.key === "Enter") {
                      e.preventDefault();
                      runSearch();
                    }
                  }}
                />
                <ChoiceGroup
                  label="Match mode"
                  selectedKey={mode}
                  options={modeOptions}
                  onChange={(_e, o) => setMode(o?.key as Mode)}
                />
                <div style={{ display: "flex", justifyContent: "flex-end" }}>
                  <PrimaryButton
                    text="Search"
                    onClick={() => runSearch()}
                    disabled={loading}
                  />
                </div>
              </Stack>
            </div>

            <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
              <div style={cardStyle}>
                <div style={{ fontWeight: 600, marginBottom: 6 }}>Source</div>
                <div
                  style={{
                    fontSize: 12,
                    color: "#605e5c",
                    marginBottom: 8,
                  }}
                >
                  Select/unselect sources. "All" toggles everything.
                </div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                  <Checkbox
                    label="All"
                    checked={allSourcesSelected}
                    onChange={handleAllSourcesToggle}
                  />
                  <Checkbox
                    label="Resource Library"
                    checked={sourceFilters.includes("resource")}
                    onChange={(_e, checked) =>
                      toggleSource("resource", !!checked)
                    }
                  />
                  <Checkbox
                    label="Weekly Words"
                    checked={sourceFilters.includes("weekly")}
                    onChange={(_e, checked) =>
                      toggleSource("weekly", !!checked)
                    }
                  />
                  <Checkbox
                    label="Lessons Learned"
                    checked={sourceFilters.includes("lessons")}
                    onChange={(_e, checked) =>
                      toggleSource("lessons", !!checked)
                    }
                  />
                  <Checkbox
                    label="IT Blogs"
                    checked={sourceFilters.includes("itblogs")}
                    onChange={(_e, checked) =>
                      toggleSource("itblogs", !!checked)
                    }
                  />
                  <Checkbox
                    label="People Directory"
                    checked={sourceFilters.includes("people")}
                    onChange={(_e, checked) =>
                      toggleSource("people", !!checked)
                    }
                  />
                  <Checkbox
                    label="Marketing Calendar"
                    checked={sourceFilters.includes("marketingcal")}
                    onChange={(_e, checked) =>
                      toggleSource("marketingcal", !!checked)
                    }
                  />
                  <Checkbox
                    label="Company Calendar"
                    checked={sourceFilters.includes("companycal")}
                    onChange={(_e, checked) =>
                      toggleSource("companycal", !!checked)
                    }
                  />
                  <Checkbox
                    label="Locations"
                    checked={sourceFilters.includes("locations")}
                    onChange={(_e, checked) =>
                      toggleSource("locations", !!checked)
                    }
                  />
                  <Checkbox
                    label="Corporate Responsibilities"
                    checked={sourceFilters.includes("corpresp")}
                    onChange={(_e, checked) =>
                      toggleSource("corpresp", !!checked)
                    }
                  />
                </div>
              </div>

              <div style={cardStyle}>
                <div style={{ fontWeight: 600, marginBottom: 6 }}>
                  File types (Resource Library)
                </div>
                <div
                  style={{
                    fontSize: 12,
                    color: "#605e5c",
                    marginBottom: 8,
                  }}
                >
                  Only applies to documents in the Resource Library.
                </div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                  {availableFileTypes.map((ext) => (
                    <Checkbox
                      key={ext}
                      label={ext.toUpperCase()}
                      checked={fileTypeFilters.includes(ext)}
                      onChange={(_e, checked) =>
                        toggleFileType(ext, !!checked)
                      }
                    />
                  ))}
                </div>
              </div>
            </div>
          </div>

          {loading && (
            <div style={{ marginTop: 20 }}>
              <Spinner label="Searching…" />
            </div>
          )}

          {error && !loading && (
            <p style={{ color: "red", marginTop: 10 }}>{error}</p>
          )}

          {!loading && paginatedResults.length > 0 && (
            <div style={{ marginTop: 20 }}>
              <div
                style={{
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "center",
                  marginBottom: 8,
                }}
              >
                <h3 style={{ margin: 0 }}>Results</h3>
                <div style={{ fontSize: 12, color: "#605e5c" }}>
                  Showing{" "}
                  <strong>
                    {(page - 1) * pageSize + 1}–
                    {Math.min(page * pageSize, results.length)}
                  </strong>{" "}
                  of <strong>{results.length}</strong>
                </div>
              </div>

              <div
                style={{
                  display: "grid",
                  gridTemplateColumns:
                    "minmax(0, 1fr) minmax(0, 1fr)",
                  gap: 12,
                }}
              >
                {paginatedResults.map((r, i) => {
                  const displayTitle =
                    r.Title || r.Filename || "(no title)";
                  const onCardClick = () => openResult(r);

                  if (r.Source === "People") {
                    const upnOrEmail =
                      r.WorkEmail ||
                      r.SipAddress ||
                      r.AccountName ||
                      "";
                    return (
                      <div
                        key={
                          (r.ListItemId || 0) +
                          "-" +
                          i +
                          "-" +
                          r.Source
                        }
                        className="vc-result-card"
                        style={{
                          ...cardStyle,
                          cursor:
                            r.WorkEmail || r.SipAddress
                              ? "pointer"
                              : "default",
                          borderLeft: "4px solid #0078d4",
                          transition: "box-shadow 0.15s ease",
                        }}
                        onClick={onCardClick}
                      >
                        <div
                          style={{
                            display: "flex",
                            justifyContent: "space-between",
                            marginBottom: 8,
                          }}
                        >
                          <span style={pillStyle}>
                            {sourceLabels[r.Source]}
                          </span>
                        </div>
                        <PeoplePersonaWithLive
                          context={context}
                          graphClient={graphClient}
                          displayName={r.PreferredName || displayTitle}
                          upnOrEmail={upnOrEmail}
                          jobTitle={r.JobTitle}
                          department={r.Department}
                        />
                        <div style={metaRowStyle}>
                          {r.WorkEmail && (
                            <span>
                              <strong>Email:</strong> {r.WorkEmail}
                            </span>
                          )}
                          {r.OfficeNumber && (
                            <span>
                              <strong>Office:</strong> {r.OfficeNumber}
                            </span>
                          )}
                          {r.CellPhone && (
                            <span>
                              <strong>Mobile:</strong> {r.CellPhone}
                            </span>
                          )}
                        </div>
                      </div>
                    );
                  }

                  const isTooltipSource =
                    r.Source === "WeeklyWords" ||
                    r.Source === "LessonsLearned" ||
                    r.Source === "ITBlogs";

                  const borderLeft =
                    r.Source === "WeeklyWords"
                      ? "4px solid #107c10"
                      : r.Source === "LessonsLearned"
                      ? "4px solid #ffb900"
                      : r.Source === "ITBlogs"
                      ? "4px solid #5c2d91"
                      : r.Source === "MarketingCalendar"
                      ? "4px solid #d83b01"
                      : r.Source === "CompanyCalendar"
                      ? "4px solid #038387"
                      : r.Source === "CorporateResponsibilities"
                      ? "4px solid #8764b8"
                      : r.Source === "Locations"
                      ? "4px solid #ca5010"
                      : "4px solid #605e5c";

                  return (
                    <div
                      key={
                        (r.ListItemId || 0) + "-" + i + "-" + r.Source
                      }
                      className="vc-result-card"
                      style={{
                        ...cardStyle,
                        cursor: r.Path ? "pointer" : "default",
                        borderLeft,
                        transition: "box-shadow 0.15s ease",
                      }}
                      onClick={onCardClick}
                    >
                      {/* Top row: badge + file type + action bar */}
                      <div
                        style={{
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                          marginBottom: 6,
                          gap: 8,
                        }}
                      >
                        <div
                          style={{
                            display: "flex",
                            alignItems: "center",
                            gap: 6,
                            minWidth: 0,
                          }}
                        >
                          <span style={pillStyle}>
                            {sourceLabels[r.Source]}
                          </span>
                          {r.FileType && (
                            <span
                              style={{
                                fontSize: 10,
                                fontWeight: 600,
                                textTransform: "uppercase",
                                color: "#605e5c",
                              }}
                            >
                              {r.FileType}
                            </span>
                          )}
                        </div>

                        {/* ✅ Action bar — visible on card hover */}
                        <CardActions
                          result={r}
                          graphClient={graphClient}
                          onCardClick={onCardClick}
                        />
                      </div>

                      {/* Title row */}
                      <div
                        style={{
                          display: "flex",
                          alignItems: "flex-start",
                          gap: 6,
                        }}
                      >
                        <div
                          style={{
                            fontWeight: 600,
                            fontSize: 14,
                            lineHeight: 1.25,
                            flex: 1,
                            minWidth: 0,
                          }}
                          dangerouslySetInnerHTML={{
                            __html: highlight(displayTitle, words),
                          }}
                        />
                        {isTooltipSource && (
                          <TooltipHost
                            styles={{ root: { display: "inline-flex" } }}
                            content={
                              <div
                                style={{
                                  maxHeight: 250,
                                  overflowY: "auto",
                                }}
                              >
                                <div
                                  dangerouslySetInnerHTML={{
                                    __html: getSafeHtmlForTooltip(
                                      r,
                                      words
                                    ),
                                  }}
                                />
                              </div>
                            }
                          >
                            <IconButton
                              iconProps={{ iconName: "Info" }}
                              title="Preview"
                              ariaLabel="Preview"
                              styles={{
                                root: {
                                  padding: 0,
                                  height: 20,
                                  width: 20,
                                  marginTop: 1,
                                },
                              }}
                              onClick={(e) => e.stopPropagation()}
                            />
                          </TooltipHost>
                        )}
                      </div>

                      {r.Source === "ResourceLibrary" &&
                        r.Filename &&
                        r.Filename !== r.Title && (
                          <div
                            style={{
                              fontSize: 12,
                              color: "#605e5c",
                              marginTop: 4,
                            }}
                          >
                            File: {r.Filename}
                          </div>
                        )}

                      {(r.Source === "MarketingCalendar" ||
                        r.Source === "CompanyCalendar") && (
                        <div style={metaRowStyle}>
                          {r.EventDate && (
                            <span>
                              <strong>Start:</strong>{" "}
                              {new Date(r.EventDate).toLocaleString()}
                            </span>
                          )}
                          {r.EndDate && (
                            <span>
                              <strong>End:</strong>{" "}
                              {new Date(r.EndDate).toLocaleString()}
                            </span>
                          )}
                          {typeof r.AllDayEvent === "boolean" && (
                            <span>
                              <strong>All day:</strong>{" "}
                              {r.AllDayEvent ? "Yes" : "No"}
                            </span>
                          )}
                          {r.EventLocation && (
                            <span>
                              <strong>Location:</strong> {r.EventLocation}
                            </span>
                          )}
                          {r.Category && (
                            <span>
                              <strong>Category:</strong> {r.Category}
                            </span>
                          )}
                          {typeof r.fRecurrence === "boolean" && (
                            <span>
                              <strong>Recurs:</strong>{" "}
                              {r.fRecurrence ? "Yes" : "No"}
                            </span>
                          )}
                        </div>
                      )}

                      {r.Source === "Locations" && (
                        <div style={metaRowStyle}>
                          {r.ProjectLocation && (
                            <span>
                              <strong>Project Location:</strong>{" "}
                              {r.ProjectLocation}
                            </span>
                          )}
                          {r.Location && (
                            <span>
                              <strong>Location:</strong> {r.Location}
                            </span>
                          )}
                          {r.Phone && (
                            <span>
                              <strong>Phone:</strong> {r.Phone}
                            </span>
                          )}
                          {r.Fax && (
                            <span>
                              <strong>Fax:</strong> {r.Fax}
                            </span>
                          )}
                        </div>
                      )}

                      {r.Source === "CorporateResponsibilities" && (
                        <div style={metaRowStyle}>
                          {r.Who && (
                            <span>
                              <strong>Who:</strong> {r.Who}
                            </span>
                          )}
                        </div>
                      )}

                      {(r.Source === "WeeklyWords" ||
                        r.Source === "LessonsLearned" ||
                        r.Source === "ITBlogs") && (
                        <div style={metaRowStyle}>
                          {r.Author && (
                            <span>
                              <strong>Author:</strong> {r.Author}
                            </span>
                          )}
                          {r.LastModifiedTime && (
                            <span>
                              <strong>Modified:</strong>{" "}
                              {new Date(
                                r.LastModifiedTime
                              ).toLocaleString()}
                            </span>
                          )}
                        </div>
                      )}

                      {r.Source === "ResourceLibrary" && (
                        <div style={metaRowStyle}>
                          {r.Author && (
                            <span>
                              <strong>Author:</strong> {r.Author}
                            </span>
                          )}
                          {r.LastModifiedTime && (
                            <span>
                              <strong>Modified:</strong>{" "}
                              {new Date(
                                r.LastModifiedTime
                              ).toLocaleString()}
                            </span>
                          )}
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>

              {/* Pagination */}
              <div
                style={{
                  marginTop: 16,
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "center",
                }}
              >
                <div>
                  <PrimaryButton
                    text="Previous"
                    disabled={page === 1}
                    onClick={() => setPage(page - 1)}
                    style={{ marginRight: 8 }}
                  />
                  <PrimaryButton
                    text="Next"
                    disabled={page === totalPages}
                    onClick={() => setPage(page + 1)}
                  />
                </div>
                <div style={{ fontSize: 12, color: "#605e5c" }}>
                  Page {page} of {totalPages || 1}
                </div>
              </div>
            </div>
          )}

          {!loading &&
            results.length === 0 &&
            query.trim() &&
            !error && (
              <p style={{ marginTop: 16 }}>No results found.</p>
            )}
        </div>
      </div>
    </>
  );
};

export default SearchDocuments;
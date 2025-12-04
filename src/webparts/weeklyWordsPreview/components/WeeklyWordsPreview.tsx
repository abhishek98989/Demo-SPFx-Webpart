import * as React from 'react';
import { useState, useEffect } from 'react';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Dropdown, IDropdownOption } from "@fluentui/react/lib/Dropdown";
import SharePointBanners from '../../../globalCommon/Banners';

interface IArticleItem {
  Id: number;
  Title: string;
  ArticleDate: string;
  Status: string;
  VaughnContent?: string;
  PublishingContact?: {
    Title: string;
  };
}



export const WeeklyWordsPostDropdown: React.FC<any> = ({ listId, context, siteUrl }) => {
  const [articles, setArticles] = useState<IArticleItem[]>([]);
  const [selectedArticle, setSelectedArticle] = useState<IArticleItem | null>(null);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | null>(null);

  const sp = siteUrl != undefined
    ? spfi(siteUrl).using(SPFx(context))
    : spfi().using(SPFx(context));

  const loadArticles = async (): Promise<void> => {
    if (!listId) {
      setLoading(false);
      setError('List ID is required.');
      setArticles([]);
      setSelectedArticle(null);
      return;
    }

    try {
      setLoading(true);
      setError(null);

      const today = new Date();
      today.setHours(23, 59, 59, 999);
      const todayISO = today.toISOString();

      // Fetch all latest approved articles (<= today) ordered by ArticleDate desc
      const items = await sp.web.lists
        .getById(listId)
        .items
        .select(
          "Id",
          "Title",
          "ArticleDate",
          "VaughnContent",
          "PublishingContact/Title",
          "OData__ModerationStatus"
        )
        .expand("PublishingContact")
        .orderBy("ArticleDate", false)(); // newest first

      if (items && items.length > 0) {
        const typedItems = items as IArticleItem[];
        setArticles(typedItems);
        setSelectedArticle(typedItems[0]); // default to latest article
        setLoading(false);
        setError(null);
      } else {
        setArticles([]);
        setSelectedArticle(null);
        setLoading(false);
        setError('No approved articles found for today or earlier dates.');
      }
    } catch (err: any) {
      console.error('Error loading articles:', err);
      setLoading(false);
      setError(`Error loading articles: ${err.message}`);
      setArticles([]);
      setSelectedArticle(null);
    }
  };

  useEffect(() => {
    loadArticles();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [listId]);

  const formatDateTime = (dateString: string): { date: string; time: string } => {
    const articleDate = new Date(dateString);

    const formattedDate = articleDate.toLocaleDateString('en-US', {
      weekday: 'long',
      year: 'numeric',
      month: 'long',
      day: 'numeric'
    });

    const formattedTime = articleDate.toLocaleTimeString('en-US', {
      hour: '2-digit',
      minute: '2-digit',
      hour12: true
    });

    return { date: formattedDate, time: formattedTime };
  };

  // Inline Styles
  const containerStyle: React.CSSProperties = {
    fontFamily: '"Segoe UI", Tahoma, Geneva, Verdana, sans-serif',
    margin: '0 auto',
    padding: '20px',
    paddingTop: '0px'
  };

  const loadingStyle: React.CSSProperties = {
    textAlign: 'center',
    padding: '40px 20px',
    color: '#666'
  };

  const spinnerStyle: React.CSSProperties = {
    width: '40px',
    height: '40px',
    margin: '0 auto 20px',
    border: '4px solid #f3f3f3',
    borderTop: '4px solid #0078d4',
    borderRadius: '50%',
    animation: 'spin 1s linear infinite'
  };

  const errorStyle: React.CSSProperties = {
    textAlign: 'center',
    padding: '40px 20px',
    color: '#d13438'
  };

  const errorTitleStyle: React.CSSProperties = {
    marginBottom: '15px',
    color: '#d13438'
  };

  const retryButtonStyle: React.CSSProperties = {
    background: '#0078d4',
    color: 'white',
    border: 'none',
    padding: '10px 20px',
    borderRadius: '4px',
    cursor: 'pointer',
    fontSize: '14px'
  };

  const noContentStyle: React.CSSProperties = {
    textAlign: 'center',
    padding: '40px 20px',
    color: '#666'
  };

  const noPageContentStyle: React.CSSProperties = {
    textAlign: 'center',
    padding: '40px 20px',
    background: '#faf9f8',
    borderRadius: '8px',
    border: '1px solid #f3f2f1',
    fontStyle: 'italic',
    color: '#605e5c'
  };

  // CSS animation keyframes
  const spinKeyframes = `
    @keyframes spin {
      0% { transform: rotate(0deg); }
      100% { transform: rotate(360deg); }
    }
  `;

  const onArticleChange = (
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (!option) return;
    const selectedId = option.key as number;
    const found = articles.find(a => a.Id === selectedId) || null;
    setSelectedArticle(found);
  };

  const dropdownOptions: IDropdownOption[] = articles.map(article => {
    let label = article.Title;
    try {
      const { date } = formatDateTime(article.ArticleDate);
      label = `${date} - ${article.Title}`;
    } catch {
      // fallback to title only if date is invalid
      label = article.Title;
    }
    return {
      key: article.Id,
      text: label
    };
  });

  if (loading) {
    return (
      <div style={containerStyle}>
        <style>{spinKeyframes}</style>
        <div style={loadingStyle}>
          <div style={spinnerStyle}></div>
          <p>Loading articles...</p>
        </div>
      </div>
    );
  }

  if (error && articles.length === 0) {
    return (
      <div style={containerStyle}>
        <div style={errorStyle}>
          <h3 style={errorTitleStyle}>⚠️ Unable to Load Content</h3>
          <p style={{ marginBottom: '20px', color: '#666' }}>{error}</p>
          <button
            style={retryButtonStyle}
            onClick={loadArticles}
            onMouseOver={(e) => e.currentTarget.style.background = '#106ebe'}
            onMouseOut={(e) => e.currentTarget.style.background = '#0078d4'}
          >
            Try Again
          </button>
        </div>
      </div>
    );
  }

  if (!selectedArticle) {
    return (
      <div style={containerStyle}>
        <div style={noContentStyle}>
          <h3 style={{ marginBottom: '15px', color: '#323130' }}>📝 No Articles Available</h3>
          <p>No approved articles found for today or earlier dates.</p>
        </div>
      </div>
    );
  }

  const { date } = formatDateTime(selectedArticle.ArticleDate);

  return (
    <>
      <SharePointBanners context={context} />
      <div style={containerStyle} className="WeeklyWordsPost">
        <h2 style={{ textAlign: "justify" }} className="ms-webpart-titleText">
          <span style={{ whiteSpace: "nowrap" }}>
            <span>Weekly Words from Westpark</span>
            <span id="WebPartCaptionWPQ5" />
          </span>
        </h2>

        {/* NEW: Dropdown to select article */}
        <div style={{ marginBottom: '16px', maxWidth: '500px' }}>
          <Dropdown
            placeholder="Select an article"
            label="Weekly Words posts"
            options={dropdownOptions}
            selectedKey={selectedArticle.Id}
            onChange={onArticleChange}
          />
        </div>

        {/* Selected article details */}
        <div>
          <div className="link-item">
            <span
              style={{
                fontWeight: 600,
                fontStyle: "italic",
                paddingRight: "2px",
                color: "#000000"
              }}
            >
              Posted by:
            </span>
            {selectedArticle?.PublishingContact?.Title}
            <br />
            <span
              style={{
                fontWeight: 600,
                fontStyle: "italic",
                paddingRight: "2px",
                color: "#000000"
              }}
            >
              Date Posted :
            </span>
            {date}
          </div>

          <div>
            {selectedArticle.VaughnContent ? (
              <div
                dangerouslySetInnerHTML={{ __html: selectedArticle.VaughnContent }}
                style={{
                  width: "100%",
                  maxWidth: "100%",
                  overflowWrap: "break-word",
                  wordBreak: "break-word"
                }}
              />
            ) : (
              <div style={noPageContentStyle}>
                <p>No page content available for this article.</p>
              </div>
            )}
          </div>
        </div>
      </div>
    </>
  );
};

export default WeeklyWordsPostDropdown;

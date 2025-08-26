import * as React from 'react';
import { useState, useEffect } from 'react';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

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

interface IWeeklyWordsProps {
    listId: string;
    siteUrl: string;
    context: any; // SharePoint context
}

export const WeeklyWordsPost: React.FC<IWeeklyWordsProps> = ({ listId, context, siteUrl }) => {
    const [article, setArticle] = useState<IArticleItem | null>(null);
    const [loading, setLoading] = useState<boolean>(true);
    const [error, setError] = useState<string | null>(null);

    const sp = siteUrl != undefined ? spfi(siteUrl).using(SPFx(context)) : spfi().using(SPFx(context));

    const loadLatestArticle = async (): Promise<void> => {
        if (!listId) {
            setLoading(false);
            setError('List ID is required.');
            setArticle(null);
            return;
        }

        try {
            setLoading(true);
            setError(null);

            const today = new Date();
            today.setHours(23, 59, 59, 999);
            const todayISO = today.toISOString();

            // Fetch the latest approved article using PnP/SP
            const items = await sp.web.lists
                .getById(listId)
                .items
                .select("Id", "Title", "ArticleDate", "VaughnContent", "PublishingContact/Title", "OData__ModerationStatus")
                .expand("PublishingContact")
                .filter(`(OData__ModerationStatus eq 0) and (ArticleDate le datetime'${todayISO}')`)
                .orderBy("ArticleDate", false)
                .top(1)();

            if (items && items.length > 0) {
                setArticle(items[0] as IArticleItem);
                setLoading(false);
                setError(null);
            } else {
                setArticle(null);
                setLoading(false);
                setError('No approved articles found for today or earlier dates.');
            }
        } catch (err: any) {
            console.error('Error loading article:', err);
            setLoading(false);
            setError(`Error loading article: ${err.message}`);
            setArticle(null);
        }
    };

    useEffect(() => {
        loadLatestArticle();
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
        maxWidth: '1000px',
        margin: '0 auto',
        padding: '20px',
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

    const articleHeaderStyle: React.CSSProperties = {
        marginBottom: '30px',
        borderBottom: '2px solid #f3f2f1',
        paddingBottom: '20px'
    };

    const titleStyle: React.CSSProperties = {
        fontSize: '28px',
        fontWeight: 600,
        color: '#323130',
        margin: '0 0 15px 0',
        lineHeight: '1.3'
    };

    const metadataStyle: React.CSSProperties = {
        display: 'flex',
        justifyContent: 'space-between',
        alignItems: 'center',
        flexWrap: 'wrap',
        gap: '15px'
    };

    const publishInfoStyle: React.CSSProperties = {
        display: 'flex',
        flexDirection: 'column',
        gap: '8px'
    };

    const dateTimeStyle: React.CSSProperties = {
        color: '#605e5c',
        fontSize: '14px',
        fontWeight: 500
    };

    const authorStyle: React.CSSProperties = {
        color: '#605e5c',
        fontSize: '14px'
    };

    const statusBadgeStyle: React.CSSProperties = {
        background: '#dff6dd',
        color: '#107c10',
        padding: '6px 12px',
        borderRadius: '16px',
        fontSize: '12px',
        fontWeight: 600
    };

    const contentStyle: React.CSSProperties = {
        fontSize: "16px",
        color: "#323130",
        marginBottom: "30px",
        lineHeight: "1.6",
        width: "100%",
        maxWidth: "100%",
        overflowWrap: "break-word",
    };

    const globalContentCSS = `
  .article-content p { margin: 0 0 1em; }
  .article-content h1, .article-content h2, .article-content h3 {
    margin: 1em 0 0.5em;
    font-weight: 600;
    line-height: 1.3;
  }
  .article-content img {
    max-width: 100%;
    height: auto;
    display: block;
    margin: 1em auto;
  }
  .article-content table {
    width: 100%;
    border-collapse: collapse;
  }
  .article-content table td, .article-content table th {
    border: 1px solid #ddd;
    padding: 8px;
  }
`;


    const noPageContentStyle: React.CSSProperties = {
        textAlign: 'center',
        padding: '40px 20px',
        background: '#faf9f8',
        borderRadius: '8px',
        border: '1px solid #edebe9',
        fontStyle: 'italic',
        color: '#605e5c'
    };

    const footerStyle: React.CSSProperties = {
        textAlign: 'right',
        paddingTop: '20px',
        borderTop: '1px solid #f3f2f1',
        color: '#a19f9d',
        fontSize: '12px'
    };

    // Add CSS animation keyframes
    const spinKeyframes = `
    @keyframes spin {
      0% { transform: rotate(0deg); }
      100% { transform: rotate(360deg); }
    }
  `;

    if (loading) {
        return (
            <div style={containerStyle}>
                <style>{spinKeyframes}</style>
                <div style={loadingStyle}>
                    <div style={spinnerStyle}></div>
                    <p>Loading latest article...</p>
                </div>
            </div>
        );
    }

    if (error) {
        return (
            <div style={containerStyle}>
                <div style={errorStyle}>
                    <h3 style={errorTitleStyle}>⚠️ Unable to Load Content</h3>
                    <p style={{ marginBottom: '20px', color: '#666' }}>{error}</p>
                    <button
                        style={retryButtonStyle}
                        onClick={loadLatestArticle}
                        onMouseOver={(e) => e.currentTarget.style.background = '#106ebe'}
                        onMouseOut={(e) => e.currentTarget.style.background = '#0078d4'}
                    >
                        Try Again
                    </button>
                </div>
            </div>
        );
    }

    if (!article) {
        return (
            <div style={containerStyle}>
                <div style={noContentStyle}>
                    <h3 style={{ marginBottom: '15px', color: '#323130' }}>📝 No Articles Available</h3>
                    <p>No approved articles found for today or earlier dates.</p>
                </div>
            </div>
        );
    }

    const { date, time } = formatDateTime(article.ArticleDate);

    return (
        <div style={containerStyle}>
            <h2 style={{ textAlign: "justify" }} className="ms-webpart-titleText">
                <span style={{ whiteSpace: "nowrap" }}>
                    <span>Weekly Words from Westpark</span>
                    <span id="WebPartCaptionWPQ5" />
                </span>
            </h2>
            <div>
                <div className="link-item">
                    <span style={{ fontWeight: 600, fontStyle: "italic", paddingRight: "2px", color: "#000000" }}>
                        Posted by:
                    </span>
                    {article?.PublishingContact?.Title}
                    <br />
                    <span style={{ fontWeight: 600, fontStyle: "italic", paddingRight: "2px", color: "#000000" }}>
                        Date Posted :
                    </span>
                    {date}
                </div>
                <style>{globalContentCSS}</style>
                <div style={contentStyle}>
                    {article.VaughnContent ? (
                        <div
                            dangerouslySetInnerHTML={{ __html: article.VaughnContent }}
                            style={{
                                ...contentStyle,
                                width: "100%",
                                maxWidth: "100%",
                                overflowWrap: "break-word",
                                wordBreak: "break-word",
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
    );
};

export default WeeklyWordsPost;
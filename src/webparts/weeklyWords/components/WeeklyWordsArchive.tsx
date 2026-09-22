import * as React from 'react';
import { useState, useEffect } from 'react';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/security";
import { PermissionKind } from "@pnp/sp/security";
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';

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

interface IWeeklyWordsArchiveProps {
    listId: string;
    siteUrl: string;
    context: any;
}

export const WeeklyWordsArchive: React.FC<IWeeklyWordsArchiveProps> = ({ listId, context, siteUrl }) => {
    const [articles, setArticles] = useState<IArticleItem[]>([]);
    const [loading, setLoading] = useState<boolean>(true);
    const [error, setError] = useState<string | null>(null);
    const [selectedArticle, setSelectedArticle] = useState<IArticleItem | null>(null);
    const [isPanelOpen, setIsPanelOpen] = useState<boolean>(false);
    const [hoveredRow, setHoveredRow] = useState<number | null>(null);
    const [canEdit, setCanEdit] = useState<boolean>(false);

    const ARCHIVE_PAGE_URL = "https://vaughnconstruction.sharepoint.com/news/SitePages/Weekly-Words.aspx";

    const sp = siteUrl != undefined ? spfi(siteUrl).using(SPFx(context)) : spfi().using(SPFx(context));

    const loadAllArticles = async (): Promise<void> => {
        if (!listId) {
            setLoading(false);
            setError('List ID is required.');
            return;
        }

        try {
            setLoading(true);
            setError(null);

            // Check if current user has edit access on the list
            const perms = await sp.web.lists
                .getById(listId)
                .effectiveBasePermissions();
            const hasEdit = sp.web.lists
                .getById(listId)
                .hasPermissions(perms, PermissionKind.EditListItems);
            setCanEdit(hasEdit);

            const today = new Date();
            today.setHours(23, 59, 59, 999);
            const todayISO = today.toISOString();

            const items = await sp.web.lists
                .getById(listId)
                .items
                .select("Id", "Title", "ArticleDate", "VaughnContent", "PublishingContact/Title", "OData__ModerationStatus")
                .expand("PublishingContact")
                .filter(`(OData__ModerationStatus eq 0) and (ArticleDate le datetime'${todayISO}')`)
                .orderBy("ArticleDate", false)
                .top(5000)();

            setArticles(items as IArticleItem[]);
            setLoading(false);
            setError(null);
        } catch (err: any) {
            console.error('Error loading articles:', err);
            setLoading(false);
            setError(`Error loading articles: ${err.message}`);
        }
    };

    useEffect(() => {
        loadAllArticles();
    }, [listId]);

    const formatDate = (dateString: string): string => {
        const date = new Date(dateString);
        return date.toLocaleDateString('en-US', {
            weekday: 'long',
            year: 'numeric',
            month: 'long',
            day: 'numeric'
        });
    };

    const formatDateShort = (dateString: string): string => {
        const date = new Date(dateString);
        return date.toLocaleDateString('en-US', {
            year: 'numeric',
            month: 'short',
            day: 'numeric'
        });
    };

    const openPanel = (article: IArticleItem): void => {
        setSelectedArticle(article);
        setIsPanelOpen(true);
    };

    const closePanel = (): void => {
        setIsPanelOpen(false);
        setSelectedArticle(null);
    };

    // ── Styles ──────────────────────────────────────────────────────────────
    const containerStyle: React.CSSProperties = {
        fontFamily: '"Segoe UI", Tahoma, Geneva, Verdana, sans-serif',
        margin: '0 auto',
        padding: '20px',
        paddingTop: '0px'
    };

    const headerStyle: React.CSSProperties = {
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'space-between',
        marginBottom: '16px',
        paddingBottom: '12px',
        borderBottom: '2px solid #0078d4'
    };

    const titleStyle: React.CSSProperties = {
        fontSize: '20px',
        fontWeight: 600,
        color: '#323130',
        margin: 0
    };

    const titleLinkStyle: React.CSSProperties = {
        fontSize: '20px',
        fontWeight: 600,
        color: '#0078d4',
        margin: 0,
        textDecoration: 'none',
        cursor: 'pointer',
        transition: 'color 0.15s ease'
    };

    const badgeStyle: React.CSSProperties = {
        background: '#0078d4',
        color: '#fff',
        borderRadius: '12px',
        padding: '2px 10px',
        fontSize: '12px',
        fontWeight: 600
    };

    const tableWrapStyle: React.CSSProperties = {
        overflowX: 'auto',
        borderRadius: '6px',
        border: '1px solid #edebe9',
        boxShadow: '0 2px 8px rgba(0,0,0,0.06)'
    };

    const tableStyle: React.CSSProperties = {
        width: '100%',
        borderCollapse: 'collapse',
        fontSize: '14px'
    };

    const theadStyle: React.CSSProperties = {
        background: '#f3f2f1'
    };

    const thStyle: React.CSSProperties = {
        padding: '12px 16px',
        textAlign: 'left',
        fontWeight: 600,
        color: '#323130',
        borderBottom: '1px solid #d2d0ce',
        whiteSpace: 'nowrap'
    };

    const getTdStyle = (isHovered: boolean): React.CSSProperties => ({
        padding: '12px 16px',
        borderBottom: '1px solid #edebe9',
        color: '#323130',
        verticalAlign: 'middle',
        background: isHovered ? '#f0f6ff' : 'transparent',
        transition: 'background 0.15s ease'
    });

    const getTrStyle = (isHovered: boolean): React.CSSProperties => ({
        cursor: 'pointer',
        transition: 'background 0.15s ease'
    });

    const titleCellStyle = (isHovered: boolean): React.CSSProperties => ({
        ...getTdStyle(isHovered),
        color: isHovered ? '#0078d4' : '#323130',
        fontWeight: 500,
        textDecoration: isHovered ? 'underline' : 'none'
    });

    const viewBtnStyle = (isHovered: boolean): React.CSSProperties => ({
        display: 'inline-block',
        padding: '4px 12px',
        borderRadius: '4px',
        border: `1px solid ${isHovered ? '#0078d4' : '#d2d0ce'}`,
        background: isHovered ? '#0078d4' : 'transparent',
        color: isHovered ? '#fff' : '#0078d4',
        fontSize: '12px',
        fontWeight: 500,
        cursor: 'pointer',
        transition: 'all 0.15s ease',
        whiteSpace: 'nowrap' as const
    });

    const spinnerWrapStyle: React.CSSProperties = {
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        padding: '60px 20px',
        flexDirection: 'column',
        gap: '12px'
    };

    const errorWrapStyle: React.CSSProperties = {
        textAlign: 'center',
        padding: '40px 20px',
        color: '#d13438'
    };

    const emptyWrapStyle: React.CSSProperties = {
        textAlign: 'center',
        padding: '60px 20px',
        color: '#605e5c',
        background: '#faf9f8',
        borderRadius: '6px',
        border: '1px solid #edebe9'
    };

    // ── Panel content styles ─────────────────────────────────────────────────
    const panelHeaderStyle: React.CSSProperties = {
        paddingBottom: '12px',
        borderBottom: '2px solid #0078d4',
        marginBottom: '16px'
    };

    const panelTitleStyle: React.CSSProperties = {
        fontSize: '20px',
        fontWeight: 700,
        color: '#323130',
        marginBottom: '6px'
    };

    const panelMetaStyle: React.CSSProperties = {
        fontSize: '13px',
        color: '#605e5c'
    };

    const panelMetaLabelStyle: React.CSSProperties = {
        fontWeight: 600,
        fontStyle: 'italic',
        color: '#000',
        marginRight: '4px'
    };

    const noPageContentStyle: React.CSSProperties = {
        textAlign: 'center',
        padding: '40px 20px',
        background: '#faf9f8',
        borderRadius: '8px',
        border: '1px solid #edebe9',
        fontStyle: 'italic',
        color: '#605e5c'
    };

    // ── Loading state ────────────────────────────────────────────────────────
    if (loading) {
        return (
            <div style={containerStyle}>
                <div style={spinnerWrapStyle}>
                    <Spinner size={SpinnerSize.large} label="Loading archive..." />
                </div>
            </div>
        );
    }

    // ── Error state ──────────────────────────────────────────────────────────
    if (error) {
        return (
            <div style={containerStyle}>
                <div style={errorWrapStyle}>
                    <h3 style={{ color: '#d13438', marginBottom: '12px' }}>⚠️ Unable to Load Archive</h3>
                    <p style={{ color: '#666', marginBottom: '20px' }}>{error}</p>
                    <button
                        style={{ background: '#0078d4', color: '#fff', border: 'none', padding: '10px 20px', borderRadius: '4px', cursor: 'pointer', fontSize: '14px' }}
                        onClick={loadAllArticles}
                    >
                        Try Again
                    </button>
                </div>
            </div>
        );
    }

    // ── Render ───────────────────────────────────────────────────────────────
    return (
        <div style={containerStyle}>
            {/* Header */}
            <div style={headerStyle}>
                {canEdit ? (
                    <h2 style={{ margin: 0 }}>
                        <a
                            href={ARCHIVE_PAGE_URL}
                            target="_blank"
                            rel="noreferrer"
                            style={titleLinkStyle}
                            onMouseEnter={(e) => { e.currentTarget.style.color = '#005a9e'; e.currentTarget.style.textDecoration = 'underline'; }}
                            onMouseLeave={(e) => { e.currentTarget.style.color = '#0078d4'; e.currentTarget.style.textDecoration = 'none'; }}
                            title="Open Weekly Words management page"
                        >
                            📚 Weekly Words – Archive ↗
                        </a>
                    </h2>
                ) : (
                    <h2 style={titleStyle}>📚 Weekly Words – Archive</h2>
                )}
                <span style={badgeStyle}>{articles.length} post{articles.length !== 1 ? 's' : ''}</span>
            </div>

            {articles.length === 0 ? (
                <div style={emptyWrapStyle}>
                    <h3 style={{ color: '#323130', marginBottom: '8px' }}>📝 No Posts Yet</h3>
                    <p>No approved articles have been published yet.</p>
                </div>
            ) : (
                <div style={tableWrapStyle}>
                    <table style={tableStyle}>
                        <thead style={theadStyle}>
                            <tr>
                                <th style={thStyle}>#</th>
                                <th style={thStyle}>Title</th>
                                <th style={thStyle}>Posted By</th>
                                <th style={thStyle}>Date Posted</th>
                                <th style={{ ...thStyle, textAlign: 'center' }}>View</th>
                            </tr>
                        </thead>
                        <tbody>
                            {articles.map((article, index) => {
                                const isHovered = hoveredRow === article.Id;
                                return (
                                    <tr
                                        key={article.Id}
                                        style={getTrStyle(isHovered)}
                                        onMouseEnter={() => setHoveredRow(article.Id)}
                                        onMouseLeave={() => setHoveredRow(null)}
                                        onClick={() => openPanel(article)}
                                    >
                                        <td style={{ ...getTdStyle(isHovered), color: '#a19f9d', width: '48px' }}>
                                            {index + 1}
                                        </td>
                                        <td style={titleCellStyle(isHovered)}>
                                            {article.Title}
                                        </td>
                                        <td style={getTdStyle(isHovered)}>
                                            {article.PublishingContact?.Title || '—'}
                                        </td>
                                        <td style={{ ...getTdStyle(isHovered), whiteSpace: 'nowrap' }}>
                                            {article.ArticleDate ? formatDateShort(article.ArticleDate) : '—'}
                                        </td>
                                        <td style={{ ...getTdStyle(isHovered), textAlign: 'center' }}>
                                            <span style={viewBtnStyle(isHovered)}>
                                                {isHovered ? 'Open →' : 'View'}
                                            </span>
                                        </td>
                                    </tr>
                                );
                            })}
                        </tbody>
                    </table>
                </div>
            )}

            {/* Detail Panel */}
            <Panel
                isOpen={isPanelOpen}
                onDismiss={closePanel}
                type={PanelType.large}
                isLightDismiss
                headerText="Weekly Words from Westpark"
                closeButtonAriaLabel="Close"
            >
                {selectedArticle && (
                    <div>
                        {/* Panel article header */}
                        <div style={panelHeaderStyle}>
                            <div style={panelTitleStyle}>{selectedArticle.Title}</div>
                            <div style={panelMetaStyle}>
                                <span style={panelMetaLabelStyle}>Posted by:</span>
                                {selectedArticle.PublishingContact?.Title || '—'}
                                &nbsp;&nbsp;
                                <span style={panelMetaLabelStyle}>Date:</span>
                                {selectedArticle.ArticleDate ? formatDate(selectedArticle.ArticleDate) : '—'}
                            </div>
                        </div>

                        {/* Panel article body */}
                        {selectedArticle.VaughnContent ? (
                            <div
                                dangerouslySetInnerHTML={{ __html: selectedArticle.VaughnContent }}
                                style={{
                                    width: '100%',
                                    maxWidth: '100%',
                                    overflowWrap: 'break-word',
                                    wordBreak: 'break-word',
                                    fontSize: '14px',
                                    lineHeight: '1.6'
                                }}
                            />
                        ) : (
                            <div style={noPageContentStyle}>
                                <p>No page content available for this article.</p>
                            </div>
                        )}
                    </div>
                )}
            </Panel>
        </div>
    );
};

export default WeeklyWordsArchive;

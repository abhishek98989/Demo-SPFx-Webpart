import * as React from "react";
import { useEffect, useState } from "react";
import { Spinner, Text } from "@fluentui/react";
import { spfi, SPFx } from "@pnp/sp/presets/all";
import DOMPurify from "dompurify";
import "../../../globalCommon/style.css";

interface IBlogPost {
  Id: number;
  Title: string;
  Author: string;
  PublishedDate: string;
  PublishedYear: number;
  PublishedDateAndMonth: string;
  PublishedTime: string;
  Url: string;
  Body: string;
  CSI_Divisions: { Id: number; Title: string }[];
  Contacts: { Id: number; Title: string }[];
}

let siteUrl: any = "";
let ListIds: any = {};

const BlogPosts = (props: any) => {
  const [post, setPost] = useState<IBlogPost | null>(null);
  const [loading, setLoading] = useState<boolean>(true);

  const mapItemToPost = (item: any): IBlogPost => {
    const dateObj = item?.PublishedDate ? new Date(item.PublishedDate) : new Date(item.Created);
    const year = dateObj.getFullYear();
    const monthDate = `${dateObj.getMonth() + 1}/${dateObj.getDate()}`;
    let hours = dateObj.getHours();
    const minutes = dateObj.getMinutes().toString().padStart(2, "0");
    const ampm = hours >= 12 ? "PM" : "AM";
    hours = hours % 12 || 12;
    const time = `${hours}:${minutes} ${ampm}`;

    return {
      Id: item.Id,
      Title: item.Title || "",
      PublishedYear: year,
      PublishedDateAndMonth: monthDate,
      PublishedTime: time,
      Author: item.Author?.Title || "",
      PublishedDate: item.PublishedDate || item.Created,
      Url: item.EncodedAbsUrl || "",
      Body: item.Body || "",
      CSI_Divisions: item["CSI_x0020_Div"] ? item["CSI_x0020_Div"] : [],
      Contacts: item.Contact ? item.Contact : [],
    };
  };

  const fetchPost = async (sp: any, lessonsLearnedId: number) => {
    try {
      const item = await sp.web.lists
        .getById(ListIds?.posts)
        .items.getById(lessonsLearnedId)
        .select(
          "Id",
          "Title",
          "Author/Title",
          "Created",
          "Body",
          "EncodedAbsUrl",
          "PublishedDate",
          "CSI_x0020_Div/Id",
          "CSI_x0020_Div/Title",
          "Contact/Id",
          "Contact/Title"
        )
        .expand("Author", "CSI_x0020_Div", "Contact")();

      const mapped = mapItemToPost(item);
      setPost(mapped);
    } catch (error) {
      console.error("Error fetching post:", error);
      setPost(null);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    siteUrl = props?.Context?.pageContext?.web?.absoluteUrl;
    ListIds = {
      posts: props?.PostsListId,
    };

    const sp = spfi().using(SPFx(props.Context));

    const params = new URLSearchParams(window.location.search);
    const lessonsLearnedId = params.get("LessonsLearnedId");

    if (lessonsLearnedId) {
      fetchPost(sp, parseInt(lessonsLearnedId, 10));
    } else {
      setLoading(false);
      setPost(null);
    }
  }, [props.Context, props?.PostsListId]);

  if (loading) {
    return (
      <div className="posts-container">
        <Spinner label="Loading blog post..." />
      </div>
    );
  }

  if (!post) {
    return (
      <div className="posts-container">
        <div className="no-results">
          <div className="no-results-title">No blog post found</div>
          <div className="no-results-subtitle">
            Make sure the URL has <code>?LessonsLearnedId=##</code> and the item exists.
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="posts-container">
      {/* Breadcrumb */}
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: "6px",
          marginBottom: "12px",
          fontSize: "14px",
        }}
      >
        <a
          href="https://vaughnconstruction.sharepoint.com/ll/"
          style={{ display: "flex", alignItems: "center", textDecoration: "none", color: "#0078d4", gap: "4px" }}
        >
          {/* Fluent UI home icon class works on classic/modern if CSS is there */}
          <i className="ms-Icon ms-Icon--Home" aria-hidden="true"></i>
          <span>Home</span>
        </a>
        <span style={{ color: "#666" }}>&gt;</span>
        <span style={{ fontWeight: 500, color: "#323130" }}>{post.Title}</span>
      </div>

      <div
        className={`blog-post ms-blog-postBox ms-shadow`}
        key={post.Id}
        data-post-id={post.Id}
      >
        <div className="ms-blog-postBoxDate">
          <div className="ms-textSmall">{post.PublishedYear}</div>
          <div className="ms-textXLarge ms-blog-dateText">{post.PublishedDateAndMonth}</div>
        </div>
        <div className="ms-blog-postBoxMargin">
          <h2 className="">
            <a
              href={`${siteUrl}/SitePages/DemoLessonsLearned.aspx?LessonsLearnedId=${post.Id}`}
              className=""
            >
              {post.Title}
            </a>
          </h2>
          <div className="ms-metadata ms-textSmall" style={{ marginBottom: 6 }}>
            <span>
              by{" "}
              <span className="ms-noWrap ms-imnSpan">
                <a className="ms-subtleLink" href="#">
                  {post.Author}
                </a>
              </span>{" "}
              at {post.PublishedTime}
            </span>
          </div>

          {post.CSI_Divisions?.length > 0 && (
            <div className="ms-textSmall" style={{ marginBottom: 6 }}>
              <Text variant="small">CSI Divisions: </Text>
              {post.CSI_Divisions.map((csi, index) => (
                <span key={csi.Id}>
                  {index > 0 && ", "}
                  {csi.Title}
                </span>
              ))}
            </div>
          )}

          {post.Contacts?.length > 0 && (
            <div className="ms-textSmall" style={{ marginBottom: 10 }}>
              <Text variant="small">Contact(s): </Text>
              {post.Contacts.map((contact, index) => (
                <span key={contact.Id}>
                  {index > 0 && ", "}
                  {contact.Title}
                </span>
              ))}
            </div>
          )}

          <div className="ms-blog-postBody">
            <div
              dir=""
              className="ms-rtestate-field"
              dangerouslySetInnerHTML={{
                __html: DOMPurify.sanitize(post.Body),
              }}
            />
          </div>
        </div>
      </div>
    </div>
  );
};

export default BlogPosts;

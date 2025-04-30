import * as React from "react";
import { useEffect, useState, useRef } from "react";
import { TextField, Stack, DefaultButton, Label, IconButton } from "@fluentui/react";
import { spfi,SPFx } from "@pnp/sp/presets/all";

interface IBlogPost {
  Id: number;
  Title: string;
  Author: string;
  PublishedDate: string;
  Categories: string[];
  Url: string;
}

const BlogPosts: React.FC = (props:any) => {
  const [posts, setPosts] = useState<any[]>([]);
  const [searchQuery, setSearchQuery] = useState<string>("");
  const [searchResults, setSearchResults] = useState<number[]>([]);
  const [currentResultIndex, setCurrentResultIndex] = useState<number>(0);
  const resultsRefs = useRef<(HTMLDivElement | null)[]>([]);

  useEffect(() => {
    const sp = spfi().using(SPFx(props.context));
    const fetchPosts = async () => {
      const items = await sp.web.lists.getById("6991af0c-d310-4590-bbd2-2e25cd59eebd").items.select("Id", "Title", "Author/Title", "Created", "Body", "EncodedAbsUrl").expand("Author")();
      const mappedPosts = items.map((item: any) => ({
        Id: item.Id,
        Title: item.Title,
        Author: item.Author?.Title || "",
        PublishedDate: new Date(item.Created).toLocaleDateString(),
        // Categories: item.Categories?.split(";") || [],
        Url: item.EncodedAbsUrl
      }));
      setPosts(mappedPosts);
    };

    fetchPosts();
  }, []);

  const handleSearch = (text: string) => {
    setSearchQuery(text);
    const matches = posts.reduce<number[]>((acc, post, idx) => {
      const content = `${post.Title} ${post.Author} ${post.PublishedDate} ${post.Categories.join(" ")}`.toLowerCase();
      if (content.includes(text.toLowerCase())) {
        acc.push(idx);
      }
      return acc;
    }, []);
    setSearchResults(matches);
    setCurrentResultIndex(0);

    if (matches.length && resultsRefs.current[matches[0]]) {
      resultsRefs.current[matches[0]]?.scrollIntoView({ behavior: "smooth", block: "center" });
    }
  };

  const highlightText = (text: string) => {
    if (!searchQuery) return text;
    const parts = text.split(new RegExp(`(${searchQuery})`, "gi"));
    return parts.map((part, index) => (
      part.toLowerCase() === searchQuery.toLowerCase() ? <mark key={index}>{part}</mark> : part
    ));
  };

  const moveToNext = () => {
    if (searchResults.length === 0) return;
    const nextIndex = (currentResultIndex + 1) % searchResults.length;
    setCurrentResultIndex(nextIndex);
    const postIndex = searchResults[nextIndex];
    resultsRefs.current[postIndex]?.scrollIntoView({ behavior: "smooth", block: "center" });
  };

  const moveToPrev = () => {
    if (searchResults.length === 0) return;
    const prevIndex = (currentResultIndex - 1 + searchResults.length) % searchResults.length;
    setCurrentResultIndex(prevIndex);
    const postIndex = searchResults[prevIndex];
    resultsRefs.current[postIndex]?.scrollIntoView({ behavior: "smooth", block: "center" });
  };

  return (
    <Stack tokens={{ childrenGap: 10 }} styles={{ root: { padding: 20 } }}>
      <TextField
        label="Search"
        value={searchQuery}
        onChange={(_, newValue) => handleSearch(newValue || "")}
        onRenderSuffix={() => (
          <>
            <IconButton iconProps={{ iconName: "Up" }} onClick={moveToPrev} disabled={searchResults.length === 0} />
            <IconButton iconProps={{ iconName: "Down" }} onClick={moveToNext} disabled={searchResults.length === 0} />
          </>
        )}
      />
      <Label>
        {searchResults.length > 0
          ? `Match ${currentResultIndex + 1} of ${searchResults.length}`
          : "No matches"}
      </Label>

      {posts.map((post, index) => {
        const isHighlighted = searchResults.includes(index);
        return (
          <div
            key={post.Id}
            ref={el => resultsRefs.current[index] = el}
            style={{
              border: "1px solid #ccc",
              borderRadius: 4,
              padding: 15,
              marginBottom: 10,
              backgroundColor: isHighlighted ? "#f5f5f5" : "white"
            }}
          >
            <div style={{ display: "flex", alignItems: "center", marginBottom: 8 }}>
              <div style={{ fontSize: "18px", fontWeight: 600 }}>
                <a href={post.Url} target="_blank" rel="noopener noreferrer" style={{ textDecoration: "none" }}>
                  {highlightText(post.Title)}
                </a>
              </div>
            </div>
            <div style={{ fontSize: "14px", color: "#666" }}>
              <span>By {highlightText(post.Author)} on {highlightText(post.PublishedDate)}</span>
            </div>
            {/* <div style={{ marginTop: 5 }}>
              Categories: {post.Categories.map((cat, idx) => (
                <span key={idx} style={{ marginRight: 5 }}>{highlightText(cat)}</span>
              ))}
            </div> */}
          </div>
        );
      })}
    </Stack>
  );
};

export default BlogPosts;

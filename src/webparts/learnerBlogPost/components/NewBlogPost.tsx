import * as React from "react";
import { useEffect, useState, useRef, useCallback } from "react";
import { TextField, Stack, Label, IconButton, Spinner } from "@fluentui/react";
import { spfi, SPFx } from "@pnp/sp/presets/all";
import DOMPurify from "dompurify";

interface IBlogPost {
  Id: number;
  Title: string;
  Author: string;
  PublishedDate: string;
  Categories: string[];
  Url: string;
  Body: string;
}

const BlogPosts = (props: any) => {
  const [allPosts, setAllPosts] = useState<IBlogPost[]>([]);
  const [visiblePosts, setVisiblePosts] = useState<IBlogPost[]>([]);
  const [searchQuery, setSearchQuery] = useState<string>("");
  const [searchResults, setSearchResults] = useState<number[]>([]);
  const [currentResultIndex, setCurrentResultIndex] = useState<number>(0);
  const [loading, setLoading] = useState<boolean>(true);
  const [page, setPage] = useState<number>(1);
  const postsPerPage = 3;
  
  const resultsRefs = useRef<(HTMLDivElement | null)[]>([]);
  const observerRef = useRef<IntersectionObserver | null>(null);
  const loaderRef = useRef<HTMLDivElement | null>(null);

  // Fetch posts from SharePoint
  useEffect(() => {
    const sp = spfi().using(SPFx(props.Context));
    const fetchPosts = async () => {
      try {
        const items = await sp.web.lists.getById("6991af0c-d310-4590-bbd2-2e25cd59eebd").items
          .select("Id", "Title", "Author/Title", "Created", "Body", "EncodedAbsUrl")
          .expand("Author")();
        
        const mappedPosts = items.map((item: any) => ({
          Id: item.Id,
          Title: item.Title || "",
          Author: item.Author?.Title || "",
          PublishedDate: new Date(item.Created).toLocaleDateString(),
          Categories: [],
          Url: item.EncodedAbsUrl || "",
          Body: item.Body || ""
        }));
        
        setAllPosts(mappedPosts);
        setVisiblePosts(mappedPosts.slice(0, postsPerPage));
        setLoading(false);
      } catch (error) {
        console.error("Error fetching posts:", error);
        setLoading(false);
      }
    };

    fetchPosts();
  }, [props.Context]);

  // Setup intersection observer for infinite scrolling
  useEffect(() => {
    const options = {
      root: null,
      rootMargin: '0px',
      threshold: 1.0
    };

    // Handler when loader element is visible
    const handleObserver = (entries: IntersectionObserverEntry[]) => {
      const [entry] = entries;
      if (entry.isIntersecting && !loading && visiblePosts.length < allPosts.length) {
        setPage(prevPage => prevPage + 1);
      }
    };

    // Create observer
    observerRef.current = new IntersectionObserver(handleObserver, options);
    
    // Observe loader element if it exists
    if (loaderRef.current) {
      observerRef.current.observe(loaderRef.current);
    }

    return () => {
      if (observerRef.current && loaderRef.current) {
        observerRef.current.unobserve(loaderRef.current);
      }
    };
  }, [loading, visiblePosts.length, allPosts.length]);

  // Load more posts when page changes - filtered by search query if present
  useEffect(() => {
    // If there's a search query, only show matching posts
    const postsToDisplay = searchQuery.trim() 
      ? allPosts.filter((_, index) => searchResults.includes(index))
      : allPosts;
    
    const endIndex = page * postsPerPage;
    if (endIndex <= postsToDisplay.length) {
      setVisiblePosts(postsToDisplay.slice(0, endIndex));
    } else if (postsToDisplay.length > 0) {
      setVisiblePosts(postsToDisplay);
    }
  }, [page, allPosts, searchQuery, searchResults]);

  // Search functionality - fixed to only search within content
  const handleSearch = (text: string) => {
    setSearchQuery(text);
    
    if (!text.trim()) {
      setSearchResults([]);
      setCurrentResultIndex(0);
      // Reset to show all posts when search is cleared
      setPage(1);
      return;
    }
    
    // Search in all posts, not just visible ones
    const matches = allPosts.reduce<number[]>((acc, post, idx) => {
      // Create a temporary div to parse HTML content
      const tempDiv = document.createElement('div');
      tempDiv.innerHTML = DOMPurify.sanitize(post.Body);
      const textContent = tempDiv.textContent || tempDiv.innerText || '';
      
      const contentToSearch = `${post.Title} ${post.Author} ${post.PublishedDate} ${textContent}`.toLowerCase();
      
      if (contentToSearch.includes(text.toLowerCase())) {
        acc.push(idx);
      }
      return acc;
    }, []);
    
    setSearchResults(matches);
    setCurrentResultIndex(matches.length > 0 ? 0 : -1);

    // Reset to page 1 when searching to show filtered results from the beginning
    setPage(1);
    
    // Wait for rendering then scroll to the first result if available
    if (matches.length > 0) {
      setTimeout(() => {
        const firstMatchIndex = matches[0];
        if (resultsRefs.current[firstMatchIndex]) {
          resultsRefs.current[firstMatchIndex]?.scrollIntoView({ behavior: "smooth", block: "center" });
        }
      }, 100);
    }
  };

  // Highlight text safely (non-HTML content)
  const highlightText = (text: string) => {
    if (!searchQuery || !text) return text;
    const regex = new RegExp(`(${searchQuery.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')})`, "gi");
    const parts = text.split(regex);
    return parts.map((part, index) =>
      regex.test(part) ? <mark key={index} style={{ backgroundColor: "#ffff00" }}>{part}</mark> : part
    );
  };

  // Highlight HTML content without affecting tags
  const highlightHtmlContent = useCallback((html: string) => {
    if (!searchQuery || !html) return html;
    
    // Create a DOM parser to work with the HTML
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');
    
    // Function to recursively process text nodes only
    const processNode = (node: Node) => {
      if (node.nodeType === Node.TEXT_NODE) {
        // Only process text nodes
        const text = node.textContent || '';
        const sanitizedQuery = searchQuery.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const regex = new RegExp(`(${sanitizedQuery})`, 'gi');
        
        if (regex.test(text)) {
          // Replace the text node with highlighted content
          const span = doc.createElement('span');
          const parts = text.split(regex);
          
          parts.forEach((part, i) => {
            if (i % 2 === 1) { // Matching parts
              const mark = doc.createElement('mark');
              mark.style.backgroundColor = '#ffff00';
              mark.textContent = part;
              span.appendChild(mark);
            } else if (part) {
              span.appendChild(doc.createTextNode(part));
            }
          });
          
          // Replace the text node with our span
          if (node.parentNode) {
            node.parentNode.replaceChild(span, node);
          }
        }
      } else if (node.nodeType === Node.ELEMENT_NODE) {
        // Process child nodes recursively
        Array.from(node.childNodes).forEach(child => {
          processNode(child);
        });
      }
    };
    
    // Process the entire document body
    Array.from(doc.body.childNodes).forEach(node => {
      processNode(node);
    });
    
    // Return the modified HTML
    return doc.body.innerHTML;
  }, [searchQuery]);

  // Navigation between search results
  const moveToNext = () => {
    if (searchResults.length === 0) return;
    const nextIndex = (currentResultIndex + 1) % searchResults.length;
    setCurrentResultIndex(nextIndex);
    const postIndex = searchResults[nextIndex];
    
    // Make sure the post is visible by loading the appropriate page
    const pageForNextResult = Math.ceil((postIndex + 1) / postsPerPage);
    setPage(pageForNextResult);
    
    // Wait for rendering then scroll
    setTimeout(() => {
      resultsRefs.current[postIndex]?.scrollIntoView({ behavior: "smooth", block: "center" });
    }, 100);
  };

  const moveToPrev = () => {
    if (searchResults.length === 0) return;
    const prevIndex = (currentResultIndex - 1 + searchResults.length) % searchResults.length;
    setCurrentResultIndex(prevIndex);
    const postIndex = searchResults[prevIndex];
    
    // Make sure the post is visible by loading the appropriate page
    const pageForPrevResult = Math.ceil((postIndex + 1) / postsPerPage);
    setPage(pageForPrevResult);
    
    // Wait for rendering then scroll
    setTimeout(() => {
      resultsRefs.current[postIndex]?.scrollIntoView({ behavior: "smooth", block: "center" });
    }, 100);
  };

  const getResultCountText = () => {
    if (!searchQuery.trim()) return `Showing ${visiblePosts.length} of ${allPosts.length} posts`;
    if (searchResults.length === 0) return "No matches found";
    return `Match ${currentResultIndex + 1} of ${searchResults.length} (${visiblePosts.length} posts shown)`;
  };

  return (
    <Stack tokens={{ childrenGap: 10 }} styles={{ root: { padding: 20 } }}>
      <Stack horizontal tokens={{ childrenGap: 10 }}>
        <TextField
          label="Search"
          value={searchQuery}
          onChange={(_, newValue) => handleSearch(newValue || "")}
          styles={{ root: { width: "100%" } }}
        />
        <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 5 }}>
          <IconButton 
            iconProps={{ iconName: "Up" }} 
            onClick={moveToPrev} 
            disabled={searchResults.length === 0}
            title="Previous result"
          />
          <IconButton 
            iconProps={{ iconName: "Down" }} 
            onClick={moveToNext} 
            disabled={searchResults.length === 0}
            title="Next result"
          />
        </Stack>
      </Stack>
      
      <Label>{getResultCountText()}</Label>

      {loading ? (
        <Spinner label="Loading blog posts..." />
      ) : visiblePosts.length === 0 && searchQuery.trim() ? (
        <div style={{ textAlign: 'center', padding: '40px' }}>
          <div style={{ fontSize: '16px', fontWeight: 600, marginBottom: '8px' }}>
            No matching blog posts found
          </div>
          <div style={{ color: '#666' }}>
            Try adjusting your search criteria
          </div>
        </div>
      ) : (
        <>
          {visiblePosts.map((post, visibleIndex) => {
            // Get the actual index in the full posts array
            const actualIndex = allPosts.findIndex(p => p.Id === post.Id);
            const isHighlighted = searchResults.includes(actualIndex);
            const isCurrentResult = searchResults[currentResultIndex] === actualIndex;
            
            return (
              <div
                key={post.Id}
                ref={el => resultsRefs.current[actualIndex] = el}
                style={{
                  border: "1px solid #ccc",
                  borderRadius: 4,
                  padding: 15,
                  marginBottom: 10,
                  backgroundColor: isCurrentResult ? "#e6f7ff" : isHighlighted ? "#f5f5f5" : "white",
                  boxShadow: isCurrentResult ? "0 0 5px rgba(0, 120, 212, 0.5)" : "none",
                  transition: "all 0.3s"
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
                
                <div style={{ marginTop: 10 }}>
                  <div 
                    dangerouslySetInnerHTML={{ 
                      __html: DOMPurify.sanitize(highlightHtmlContent(post.Body)) 
                    }} 
                  />
                </div>
              </div>
            );
          })}
          
          {/* Loader element for infinite scrolling */}
          {searchQuery.trim() 
            ? (visiblePosts.length < searchResults.length && (
                <div 
                  ref={loaderRef}
                  style={{ textAlign: 'center', padding: '20px' }}
                >
                  <Spinner label="Loading more matching posts..." />
                </div>
              ))
            : (visiblePosts.length < allPosts.length && (
                <div 
                  ref={loaderRef}
                  style={{ textAlign: 'center', padding: '20px' }}
                >
                  <Spinner label="Loading more posts..." />
                </div>
              ))
          }
        </>
      )}
    </Stack>
  );
};

export default BlogPosts;
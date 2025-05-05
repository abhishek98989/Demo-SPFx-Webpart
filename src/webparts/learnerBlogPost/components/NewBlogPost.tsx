import * as React from "react";
import { useEffect, useState, useRef, useCallback } from "react";
import { TextField, Stack, Label, IconButton, Spinner, Text } from "@fluentui/react";
import { spfi, SPFx } from "@pnp/sp/presets/all";
import DOMPurify from "dompurify";
import '../../../globalCommon/style.css';
import $ from 'jquery';

interface IBlogPost {
  Id: number;
  Title: string;
  Author: string;
  PublishedDate: string;
  PublishedYear: number;
  PublishedDateAndMonth: string;
  PublishedTime: string;
  Categories: any[];
  Url: string;
  Body: string;
}

interface ICategory {
  Id: number;
  Title: string;
}

let siteUrl: any = '';
let ListIds: any = {};

const BlogPosts = (props: any) => {
  const [allPosts, setAllPosts] = useState<IBlogPost[]>([]);
  const [categories, setCategories] = useState<ICategory[]>([]);
  const [visiblePosts, setVisiblePosts] = useState<IBlogPost[]>([]);
  const [searchQuery, setSearchQuery] = useState<string>("");
  const [searchResults, setSearchResults] = useState<number[]>([]);
  const [currentResultIndex, setCurrentResultIndex] = useState<number>(0);
  const [loading, setLoading] = useState<boolean>(true);
  const [page, setPage] = useState<number>(1);
  const [selectedCategories, setSelectedCategories] = useState<ICategory[]>([]);
  const postsPerPage = 3;
  const resultsRefs = useRef<(HTMLDivElement | null)[]>([]);
  const observerRef = useRef<IntersectionObserver | null>(null);
  const loaderRef = useRef<HTMLDivElement | null>(null);
  const contentRef = useRef<HTMLDivElement | null>(null);

  // Helper function to sort posts by date (newest first)
  const sortByDate = (array: IBlogPost[]) => {  
    return [...array].sort((a, b) => new Date(b?.PublishedDate).getTime() - new Date(a?.PublishedDate).getTime());
  };

  // Fetch posts from SharePoint
  const fetchPosts = async (sp: any) => {
    try {
      const params = new URLSearchParams(window.location.search);
      let postFilter: any = params.get("PostId");
      let categoryFilter: any = params.get("CategoryId");
      const items = await sp.web.lists.getById(ListIds?.posts).items
        .select("Id", "Title", "Author/Title", "Created", "Body", "EncodedAbsUrl", "PostCategory/Title", "PostCategory/Id", "PublishedDate", "Contact/Title", "Contact/Id")
        .expand("Author,PostCategory,Contact")();

      const mappedPosts = items.map((item: any) => {
        const dateObj = new Date(item?.PublishedDate);
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
          PublishedDate: item.PublishedDate,
          Categories: item.PostCategory ? item.PostCategory : [],
          Url: item.EncodedAbsUrl || "",
          Body: item.Body || "",
        }
      });
      
      // Always ensure posts are sorted by date
      const sortedPosts = sortByDate(mappedPosts);
      setAllPosts(sortedPosts);
      
      // Apply initial filters from URL parameters
      if (postFilter != null && postFilter != undefined && postFilter != "") {
        const filteredPosts = sortedPosts?.filter((post: any) => post.Id == parseInt(postFilter));
        setVisiblePosts(filteredPosts.slice(0, postsPerPage));
      } else if (categoryFilter != null && categoryFilter != undefined && categoryFilter != "") {
        const categoryIds = categoryFilter.split(',').map((id: string) => parseInt(id));
        const filteredPosts = sortedPosts?.filter((post: any) => 
          post?.Categories?.some((cat: any) => categoryIds.includes(cat.Id))
        );
        setVisiblePosts(filteredPosts.slice(0, postsPerPage));
        
        // Set selected categories if found
        const categoriesToSelect = categories.filter(cat => categoryIds.includes(cat.Id));
        if (categoriesToSelect.length > 0) {
          setSelectedCategories(categoriesToSelect);
        }
      } else {
        setVisiblePosts(sortedPosts?.slice(0, postsPerPage));
      }
      
      setLoading(false);
    } catch (error) {
      console.error("Error fetching posts:", error);
      setLoading(false);
    }
  };

  // Fetch categories from SharePoint
  const fetchCategories = async (sp: any) => {
    try {
      const items = await sp.web.lists.getById(ListIds?.categories).items
        .select("Id", "Title")();

      setCategories(items);
    } catch (error) {
      console.error("Error fetching categories:", error);
    }
  };

  useEffect(() => {
    siteUrl = props?.Context?.pageContext?.web?.absoluteUrl;
    ListIds = {
      posts: props?.PostsListId,
      categories: props?.CategoriesListId,
    }
    const sp = spfi().using(SPFx(props.Context));
    fetchCategories(sp);
    fetchPosts(sp);
  }, [props.Context]);

  // Handle category selection - toggle selection
  const handleCategorySelect = (category: ICategory) => {
    setSelectedCategories(prevSelected => {
      const isAlreadySelected = prevSelected.some(cat => cat.Id === category.Id);
      
      if (isAlreadySelected) {
        // Remove from selection
        return prevSelected.filter(cat => cat.Id !== category.Id);
      } else {
        // Add to selection
        return [...prevSelected, category];
      }
    });
    
    // Reset search and pagination when category changes
    setSearchQuery("");
    setSearchResults([]);
    setCurrentResultIndex(0);
    setPage(1);
  };

  // Remove a specific category from selection
  const removeCategory = (categoryId: number) => {
    setSelectedCategories(prevSelected => 
      prevSelected.filter(cat => cat.Id !== categoryId)
    );
    setPage(1);
  };

  // Clear all selected categories
  const clearAllCategories = () => {
    setSelectedCategories([]);
    setPage(1);
  };

  // Get filtered posts based on current category filters
  const getFilteredPostsByCategories = useCallback(() => {
    let filteredPosts = [...allPosts];
    
    // Filter by categories if any are selected (OR logic)
    if (selectedCategories.length > 0) {
      const selectedCategoryIds = selectedCategories.map(cat => cat.Id);
      filteredPosts = filteredPosts.filter(post => 
        post.Categories.some(cat => selectedCategoryIds.includes(cat.Id))
      );
    }
    
    // Always ensure posts are sorted by date
    return sortByDate(filteredPosts);
  }, [allPosts, selectedCategories]);

  // Further filter the category-filtered posts by search query
  const getFilteredPosts = useCallback(() => {
    // First filter by categories
    let filteredPosts = getFilteredPostsByCategories();
    
    // Then apply search filter if there's a search query
    if (searchQuery.trim()) {
      // Create a temporary div to parse HTML content for each post
      filteredPosts = filteredPosts.filter(post => {
        const tempDiv = document.createElement('div');
        tempDiv.innerHTML = DOMPurify.sanitize(post.Body);
        const textContent = tempDiv.textContent || tempDiv.innerText || '';

        // Extract category titles to make them searchable
        const categoryTitles = post.Categories?.map((cat: any) => cat.Title || '').join(' ') || '';

        // Combine all searchable content including categories
        const contentToSearch = `${post.Title} ${post.Author} ${post.PublishedDate} ${post.PublishedYear} ${post.PublishedDateAndMonth} ${post.PublishedTime} ${categoryTitles} ${textContent}`.toLowerCase();

        return contentToSearch.includes(searchQuery.toLowerCase());
      });
    }
    
    // Always ensure posts are sorted by date
    return sortByDate(filteredPosts);
  }, [getFilteredPostsByCategories, searchQuery]);

  // Recalculate search results indices when search query changes
  const updateSearchResultIndices = useCallback(() => {
    if (!searchQuery.trim()) {
      setSearchResults([]);
      setCurrentResultIndex(0);
      return;
    }

    // Get the category-filtered posts first
    const categoryFilteredPosts = getFilteredPostsByCategories();
    
    // Find indices of all posts that match the search query within category-filtered posts
    const matches: number[] = [];
    
    categoryFilteredPosts.forEach((post, index) => {
      // Create a temporary div to parse HTML content
      const tempDiv = document.createElement('div');
      tempDiv.innerHTML = DOMPurify.sanitize(post.Body);
      const textContent = tempDiv.textContent || tempDiv.innerText || '';

      // Extract category titles to make them searchable
      const categoryTitles = post.Categories?.map((cat: any) => cat.Title || '').join(' ') || '';

      // Combine all searchable content including categories
      const contentToSearch = `${post.Title} ${post.Author} ${post.PublishedDate} ${post.PublishedYear} ${post.PublishedDateAndMonth} ${post.PublishedTime} ${categoryTitles} ${textContent}`.toLowerCase();

      if (contentToSearch.includes(searchQuery.toLowerCase())) {
        matches.push(index);
      }
    });

    setSearchResults(matches);
    setCurrentResultIndex(matches.length > 0 ? 0 : -1);
  }, [searchQuery, getFilteredPostsByCategories]);

  // Apply all filters and update visible posts
  const applyFiltersAndUpdatePosts = useCallback(() => {
    const filteredPosts = getFilteredPosts();
    
    // Update search results indices for highlighting
    if (searchQuery.trim()) {
      updateSearchResultIndices();
    }
    
    // Update visible posts based on current page
    const endIndex = page * postsPerPage;
    if (endIndex <= filteredPosts.length) {
      setVisiblePosts(filteredPosts.slice(0, endIndex));
    } else if (filteredPosts.length > 0) {
      setVisiblePosts(filteredPosts);
    } else {
      setVisiblePosts([]);
    }
  }, [getFilteredPosts, updateSearchResultIndices, page, postsPerPage, searchQuery]);

  // Effect for updating filtered posts when categories or search changes
  useEffect(() => {
    applyFiltersAndUpdatePosts();
  }, [selectedCategories, searchQuery, allPosts, page, applyFiltersAndUpdatePosts]);

  // Setup intersection observer for infinite scrolling
  useEffect(() => {
    // Handler when loader element is visible
    const handleObserver = (entries: IntersectionObserverEntry[]) => {
      const [entry] = entries;
      if (entry.isIntersecting && !loading) {
        const filteredPosts = getFilteredPosts();
        
        // Only load more if there are more posts to show
        if (visiblePosts.length < filteredPosts.length) {
          setPage(prevPage => prevPage + 1);
        }
      }
    };

    // Create observer with null root to use viewport
    const options = {
      root: null, // Use the browser viewport
      rootMargin: '100px',
      threshold: 0.1
    };

    // Clean up previous observer if it exists
    if (observerRef.current) {
      observerRef.current.disconnect();
    }

    observerRef.current = new IntersectionObserver(handleObserver, options);

    // Observe loader element if it exists
    if (loaderRef.current) {
      observerRef.current.observe(loaderRef.current);
    }

    return () => {
      if (observerRef.current) {
        observerRef.current.disconnect();
      }
    };
  }, [loading, visiblePosts.length, getFilteredPosts, loaderRef.current]);

  // Search functionality - now works with category filters
  const handleSearch = (text: string) => {
    setSearchQuery(text);
    setPage(1); // Reset to first page when search changes
    
    if (!text.trim()) {
      setSearchResults([]);
      setCurrentResultIndex(0);
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
    if (searchResults.length === 0 || currentResultIndex === searchResults.length - 1) return;
    const nextIndex = currentResultIndex + 1;
    setCurrentResultIndex(nextIndex);
    
    // Get the category-filtered posts
    const categoryFilteredPosts = getFilteredPostsByCategories();
    const postIndex = searchResults[nextIndex];
    const postId = categoryFilteredPosts[postIndex]?.Id;

    // Find the post in visiblePosts
    const visiblePostIndex = visiblePosts.findIndex(p => p.Id === postId);
    
    // If post is not visible, load the appropriate page
    if (visiblePostIndex === -1) {
      const pageForNextResult = Math.ceil((postIndex + 1) / postsPerPage);
      setPage(Math.max(pageForNextResult, page)); // Don't go back, only forward
    }

    // Wait for rendering then scroll to the post
    setTimeout(() => {
      // Find the post's DOM reference by its ID
      const postElement = document.querySelector(`[data-post-id="${postId}"]`) as HTMLDivElement;
      if (postElement) {
        postElement.scrollIntoView({ behavior: "smooth", block: "center" });
      }
    }, 100);
  };

  const moveToPrev = () => {
    if (searchResults.length === 0 || currentResultIndex === 0) return;
    const prevIndex = currentResultIndex - 1;
    setCurrentResultIndex(prevIndex);
    
    // Get the category-filtered posts
    const categoryFilteredPosts = getFilteredPostsByCategories();
    const postIndex = searchResults[prevIndex];
    const postId = categoryFilteredPosts[postIndex]?.Id;

    // Find the post in visiblePosts
    const visiblePostIndex = visiblePosts.findIndex(p => p.Id === postId);
    
    // If post is not visible, load the appropriate page
    if (visiblePostIndex === -1) {
      const pageForPrevResult = Math.ceil((postIndex + 1) / postsPerPage);
      setPage(pageForPrevResult);
    }

    // Wait for rendering then scroll to the post
    setTimeout(() => {
      // Find the post's DOM reference by its ID
      const postElement = document.querySelector(`[data-post-id="${postId}"]`) as HTMLDivElement;
      if (postElement) {
        postElement.scrollIntoView({ behavior: "smooth", block: "center" });
      }
    }, 100);
  };

  const getResultCountText = () => {
    const filteredPosts = getFilteredPosts();
    
    if (searchQuery.trim()) {
      return searchResults.length === 0
        ? "No matches found"
        : `Match ${currentResultIndex + 1} of ${searchResults.length}`;
    }
    
    return `Showing ${visiblePosts.length} of ${filteredPosts.length} posts`;
  };

  // Calculate if we should show the loader
  const shouldShowLoader = () => {
    const filteredPosts = getFilteredPosts();
    return visiblePosts.length < filteredPosts.length;
  };

  // Highlight category text
  const highlightCategoryText = (category: any) => {
    if (!category || !category.Title) return null;
    return highlightText(category.Title);
  };

  // Check if a category is selected
  const isCategorySelected = (categoryId: number) => {
    return selectedCategories.some(cat => cat.Id === categoryId);
  };

  // Check if a post is currently highlighted in search results
  const isHighlightedPost = (post: IBlogPost) => {
    if (!searchQuery.trim() || searchResults.length === 0) return false;
    
    const categoryFilteredPosts = getFilteredPostsByCategories();
    const postIndex = categoryFilteredPosts.findIndex(p => p.Id === post.Id);
    return searchResults.includes(postIndex);
  };

  // Check if a post is the current search result
  const isCurrentResultPost = (post: IBlogPost) => {
    if (!searchQuery.trim() || searchResults.length === 0 || currentResultIndex < 0) return false;
    
    const categoryFilteredPosts = getFilteredPostsByCategories();
    const postIndex = categoryFilteredPosts.findIndex(p => p.Id === post.Id);
    return searchResults[currentResultIndex] === postIndex;
  };

  return (
    <div className="container">
      <div className="categories-container">
        <div className="ms-core-listMenu-verticalBox ms-noList">
          <ul className="static ms-blog-listMenu-root ms-core-listMenu-root root">
            <li>
              <span className="ms-core-listMenu-item ms-blog-quickLinksTitle">Categories</span>
              <ul className="static">
                {
                  categories?.map((category: ICategory) => (
                    <li className="static" key={category.Id}>
                      <a 
                        className={`ms-link ${isCategorySelected(category.Id) ? 'ms-bold' : ''}`}
                        style={{ 
                          cursor: 'pointer',
                          fontWeight: isCategorySelected(category.Id) ? 'bold' : 'normal',
                          textDecoration: isCategorySelected(category.Id) ? 'underline' : 'none'
                        }} 
                        onClick={() => handleCategorySelect(category)}
                      >
                        {category?.Title}
                      </a>
                    </li>
                  ))
                }
              </ul>
              <ul className="static">
                <li className="static">
                  <a 
                    href="https://vaughnconstruction.sharepoint.com/ll/_layouts/15/listform.aspx?PageType=8&amp;ListId=%7B523B46CB%2D2D41%2D46D2%2D963C%2D996BE15E0495%7D&amp;RootFolder=&amp;Source=https%3A%2F%2Fvaughnconstruction%2Esharepoint%2Ecom%2Fll%2Fdefault%2Easpx" 
                    className="ms-core-listMenu-item ms-commandLink ms-uppercase ms-blog-quickLinksEntry"
                  >
                    Add Category
                  </a>
                </li>
              </ul>
            </li>
          </ul>
        </div>
      </div>
      <div className="posts-container">
        <div className="sticky-header" role="search">
          {selectedCategories.length > 0 && (
            <div style={{ 
              marginBottom: '10px', 
              padding: '5px 10px',
              backgroundColor: '#f0f0f0',
              borderRadius: '4px'
            }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '5px' }}>
                <Text variant="medium">Selected Categories:</Text>
                <IconButton
                  iconProps={{ iconName: "Clear" }}
                  onClick={clearAllCategories}
                  title="Clear all categories"
                  styles={{ root: { height: '28px' } }}
                />
              </div>
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: '5px' }}>
                {selectedCategories.map(category => (
                  <div 
                    key={category.Id} 
                    style={{
                      display: 'flex',
                      alignItems: 'center',
                      padding: '3px 8px',
                      backgroundColor: '#e0e0e0',
                      borderRadius: '3px',
                      marginBottom: '3px'
                    }}
                  >
                    <Text variant="smallPlus">{category.Title}</Text>
                    <IconButton
                      iconProps={{ iconName: "Cancel" }}
                      onClick={() => removeCategory(category.Id)}
                      title={`Remove ${category.Title}`}
                      styles={{ 
                        root: { 
                          height: '24px', 
                          width: '24px', 
                          marginLeft: '4px',
                          padding: '0'
                        },
                        icon: { fontSize: '10px' }
                      }}
                    />
                  </div>
                ))}
              </div>
            </div>
          )}
          
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
                disabled={searchResults.length === 0 || currentResultIndex === 0}
                title="Previous result"
              />
              <IconButton
                iconProps={{ iconName: "Down" }}
                onClick={moveToNext}
                disabled={searchResults.length === 0 || currentResultIndex === searchResults.length - 1}
                title="Next result"
              />
            </Stack>
          </Stack>

          <Label>{getResultCountText()}</Label>
        </div>

        <div ref={contentRef} className="scrollable-content">
          {loading ? (
            <Spinner label="Loading blog posts..." />
          ) : visiblePosts.length === 0 ? (
            <div className="no-results">
              <div className="no-results-title">No matching blog posts found</div>
              <div className="no-results-subtitle">Try adjusting your search criteria or category filters</div>
            </div>
          ) : (
            <>
              {visiblePosts.map((post: IBlogPost) => {
                const isHighlighted = isHighlightedPost(post);
                const isCurrentResult = isCurrentResultPost(post);

                return (
                  <div
                    className={`blog-post ms-blog-postBox ms-shadow ${isCurrentResult ? "current" : isHighlighted ? "highlighted" : ""}`}
                    key={post.Id}
                    data-post-id={post.Id}
                  >
                    <div className="ms-blog-postBoxDate">
                      <div className="ms-textSmall">{highlightText(post?.PublishedYear?.toString())}</div>
                      <div className="ms-textXLarge ms-blog-dateText">{highlightText(post?.PublishedDateAndMonth)}</div>
                    </div>
                    <div className="ms-blog-postBoxMargin">
                      <h2 className="">
                        <a href={`${siteUrl}/SitePages/DemoLessonsLearned.aspx?PostId=${post.Id}`} className="">{highlightText(post?.Title)}</a>
                      </h2>
                      <div className="ms-metadata ms-textSmall">
                        <span>by{' '}
                          <span className="ms-noWrap ms-imnSpan">
                            <a className="ms-subtleLink" href="/ll/_layouts/15/userdisp.aspx?ID=1679">{highlightText(post?.Author)}</a>
                          </span> at {highlightText(post?.PublishedTime)} in{' '}
                          {post?.Categories?.map((category: any, index: number) => (
                            <span key={category.Id}>
                              {index > 0 && ", "}
                              <a
                                style={{ 
                                  cursor: 'pointer',
                                  fontWeight: isCategorySelected(category.Id) ? 'bold' : 'normal',
                                  textDecoration: isCategorySelected(category.Id) ? 'underline' : 'none'
                                }}
                                onClick={() => handleCategorySelect(category)}
                                className="ms-link"
                              >
                                {highlightCategoryText(category)}
                              </a>
                            </span>
                          ))}
                        </span>
                      </div>
                      <p></p>
                      <div className="ms-blog-postBody">
                        <div
                          dir=""
                          className="ms-rtestate-field"
                          dangerouslySetInnerHTML={{ __html: DOMPurify.sanitize(highlightHtmlContent(post.Body)) }}
                        />
                      </div>
                    </div>
                  </div>
                );
              })}

              {/* This is the loader element that triggers infinite scrolling */}
              {shouldShowLoader() && (
                <div ref={loaderRef} className="loader">
                  <Spinner label="Loading more posts..." />
                </div>
              )}
            </>
          )}
        </div>
      </div>
    </div>
  );
};

export default BlogPosts;
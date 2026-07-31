export function processEmailMailtoLinks(content: string): string {
  if (!content || typeof content !== 'string') {
    return content || '';
  }

  const emailRegex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;

  if (!content.includes('@') || !emailRegex.test(content)) {
    return content;
  }
  emailRegex.lastIndex = 0;

  if (typeof window === 'undefined' || typeof DOMParser === 'undefined') {
    return content.replace(emailRegex, (email) => `<a href="mailto:${email}">${email}</a>`);
  }

  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(content, 'text/html');
    const body = doc.body;

    // 1. Process all existing <a> tags
    const anchors = Array.from(body.querySelectorAll('a'));
    anchors.forEach((a) => {
      const href = a.getAttribute('href') || '';
      const text = a.textContent || '';

      emailRegex.lastIndex = 0;
      const hrefMatch = href.match(emailRegex);
      emailRegex.lastIndex = 0;
      const textMatch = text.match(emailRegex);

      const foundEmail = hrefMatch?.[0] || textMatch?.[0];
      if (foundEmail) {
        a.setAttribute('href', `mailto:${foundEmail}`);
      }
    });

    // 2. Process text nodes (excluding nodes inside <a> tags)
    const walkTextNodes = (node: Node) => {
      if (node.nodeType === Node.TEXT_NODE) {
        const text = node.nodeValue || '';
        emailRegex.lastIndex = 0;
        if (emailRegex.test(text)) {
          emailRegex.lastIndex = 0;
          const fragment = doc.createDocumentFragment();
          let lastIndex = 0;
          let match: RegExpExecArray | null;

          while ((match = emailRegex.exec(text)) !== null) {
            const email = match[0];
            const matchIndex = match.index;

            if (matchIndex > lastIndex) {
              fragment.appendChild(doc.createTextNode(text.substring(lastIndex, matchIndex)));
            }

            const anchor = doc.createElement('a');
            anchor.setAttribute('href', `mailto:${email}`);
            anchor.textContent = email;
            fragment.appendChild(anchor);

            lastIndex = matchIndex + email.length;
          }

          if (lastIndex < text.length) {
            fragment.appendChild(doc.createTextNode(text.substring(lastIndex)));
          }

          node.parentNode?.replaceChild(fragment, node);
        }
      } else if (node.nodeType === Node.ELEMENT_NODE) {
        const elem = node as HTMLElement;
        if (elem.tagName.toLowerCase() !== 'a') {
          Array.from(elem.childNodes).forEach((child) => walkTextNodes(child));
        }
      }
    };

    Array.from(body.childNodes).forEach((child) => walkTextNodes(child));

    return body.innerHTML;
  } catch (error) {
    console.error('Error processing email mailto links:', error);
    return content.replace(emailRegex, (email) => `<a href="mailto:${email}">${email}</a>`);
  }
}

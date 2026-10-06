import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";

export const Markdown = ({ children }: { children: string }) => (
  <div className="markdown">
    <ReactMarkdown remarkPlugins={[remarkGfm]} components={{ a: (props) => <a {...props} target="_blank" rel="noreferrer" /> }}>
      {children}
    </ReactMarkdown>
  </div>
);

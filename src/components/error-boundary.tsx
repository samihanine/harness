/**
 * Keeps a failure of one part of the page from blanking the whole app: the part shows `fallback`
 * (given the error), the rest keeps working. Used around the screens, and around what is drawn from
 * text the AI wrote (markdown, diagrams), where any input can reach a rendering bug.
 */
import { Component, type ErrorInfo, type ReactNode } from "react";

type Props = { fallback: (error: Error, reset: () => void) => ReactNode; children: ReactNode; /** A change of this value tries again (e.g. the text drawn). */ resetKey?: unknown };

export class ErrorBoundary extends Component<Props, { error?: Error; key?: unknown }> {
  state: { error?: Error; key?: unknown } = { key: this.props.resetKey };
  static getDerivedStateFromError(error: Error) {
    return { error };
  }
  static getDerivedStateFromProps(props: Props, state: { error?: Error; key?: unknown }) {
    return props.resetKey !== state.key ? { error: undefined, key: props.resetKey } : null;
  }
  componentDidCatch(error: Error, info: ErrorInfo) {
    console.error("A part of the page failed to draw:", error, info.componentStack);
  }
  render() {
    return this.state.error ? this.props.fallback(this.state.error, () => this.setState({ error: undefined })) : this.props.children;
  }
}

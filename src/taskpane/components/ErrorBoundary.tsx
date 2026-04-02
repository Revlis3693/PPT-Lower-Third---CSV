import React, { Component, type ErrorInfo, type ReactNode } from "react";

type Props = { children: ReactNode };
type State = { error: Error | null };

/**
 * Surfaces real React errors in the task pane (Office WebView often hides details in the console).
 */
export class ErrorBoundary extends Component<Props, State> {
  state: State = { error: null };

  static getDerivedStateFromError(error: Error): State {
    return { error };
  }

  componentDidCatch(error: Error, info: ErrorInfo): void {
    // eslint-disable-next-line no-console
    console.error("Lower Third Builder:", error, info.componentStack);
  }

  render(): ReactNode {
    if (this.state.error) {
      return (
        <div style={{ padding: 12, fontFamily: "system-ui, sans-serif", fontSize: 13 }}>
          <h1 style={{ fontSize: 16, margin: "0 0 8px 0" }}>Something went wrong</h1>
          <pre
            style={{
              whiteSpace: "pre-wrap",
              background: "#f5f5f5",
              padding: 8,
              borderRadius: 6,
              fontSize: 12
            }}
          >
            {this.state.error.message}
          </pre>
          <p style={{ color: "#555", marginTop: 8 }}>Reload the task pane (close and reopen the add-in) after fixing the issue.</p>
        </div>
      );
    }
    return this.props.children;
  }
}

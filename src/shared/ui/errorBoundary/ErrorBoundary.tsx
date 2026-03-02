import { Component } from "react";
import type { ErrorInfo, ReactNode } from "react";
import { Button } from "@/shared/ui/button";
import { AlertTriangle } from "lucide-react";

interface Props {
  children: ReactNode;
  fallback?: ReactNode;
}

interface State {
  hasError: boolean;
  error: Error | null;
}

export class ErrorBoundary extends Component<Props, State> {
  constructor(props: Props) {
    super(props);
    this.state = { hasError: false, error: null };
  }

  static getDerivedStateFromError(error: Error): State {
    return { hasError: true, error };
  }

  componentDidCatch(error: Error, errorInfo: ErrorInfo) {
    console.error("[ErrorBoundary]", error, errorInfo);
  }

  handleReset = () => {
    this.setState({ hasError: false, error: null });
  };

  render() {
    if (this.state.hasError) {
      if (this.props.fallback) return this.props.fallback;

      return (
        <div className="flex min-h-[400px] flex-col items-center justify-center gap-4 p-8">
          <AlertTriangle className="h-16 w-16 text-red-400" />
          <div className="text-center">
            <h2 className="text-lg font-semibold text-gray-800">
              Что-то пошло не так
            </h2>
            <p className="mt-1 text-sm text-gray-500">
              {this.state.error?.message ?? "Неизвестная ошибка"}
            </p>
          </div>
          <Button variant="outline" onClick={this.handleReset}>
            Попробовать снова
          </Button>
        </div>
      );
    }

    return this.props.children;
  }
}

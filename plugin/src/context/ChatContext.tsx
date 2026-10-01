import { createContext, useContext, useState, type ReactNode } from "react";
import { useChatConversation } from "../hooks/useChatConversation";
import { useImageAttachments } from "../hooks/useImageAttachments";
import { useModelSelection } from "../hooks/useModelSelection";
import { useSuggestions } from "../hooks/useSuggestions";
import { useWordSelection } from "../onlyoffice/useWordSelection";

type ChatContextValue = {
  composer: {
    instruction: string;
    setInstruction: (value: string) => void;
    selectedContent: string;
    clearSelectedContent: () => void;
    images: ReturnType<typeof useImageAttachments>["images"];
    imageError: ReturnType<typeof useImageAttachments>["imageError"];
    imagesLoading: boolean;
    addPastedImages: (files: File[]) => Promise<void>;
    removeImage: (id: string) => void;
  };
  conversation: ReturnType<typeof useChatConversation>;
  suggestions: ReturnType<typeof useSuggestions>;
  models: ReturnType<typeof useModelSelection>;
  settings: {
    settingsOpen: boolean;
    author: string;
    setAuthor: (value: string) => void;
    openSettings: () => void;
    closeSettings: () => void;
  };
};

const ChatContext = createContext<ChatContextValue | undefined>(undefined);

export function ChatProvider({ children }: { children: ReactNode }) {
  const [instruction, setInstruction] = useState("");
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [author, setAuthor] = useState("Vespper Agent");
  const { selectedContent, clearSelectedContent } = useWordSelection();
  const imageAttachments = useImageAttachments();
  const models = useModelSelection();
  const suggestions = useSuggestions(author.trim() || "Vespper Agent");

  const conversation = useChatConversation({
    instruction,
    selectedContent,
    images: imageAttachments.images,
    imagesLoading: imageAttachments.imagesLoading,
    author,
    model: models.model,
    clearComposer: () => {
      setInstruction("");
      imageAttachments.clearImages();
    },
    applying: suggestions.applying,
    onSuggestionsProposed: suggestions.addSet,
    onSuggestionReady: suggestions.addSuggestion,
    getSuggestionReview: suggestions.getReview,
  });

  return (
    <ChatContext.Provider
      value={{
        composer: {
          instruction,
          setInstruction,
          selectedContent,
          clearSelectedContent,
          images: imageAttachments.images,
          imageError: imageAttachments.imageError,
          imagesLoading: imageAttachments.imagesLoading,
          addPastedImages: imageAttachments.addPastedImages,
          removeImage: imageAttachments.removeImage,
        },
        conversation,
        suggestions,
        models,
        settings: {
          settingsOpen,
          author,
          setAuthor,
          openSettings: () => setSettingsOpen(true),
          closeSettings: () => setSettingsOpen(false),
        },
      }}
    >
      {children}
    </ChatContext.Provider>
  );
}

function useChatContext(): ChatContextValue {
  const context = useContext(ChatContext);
  if (!context) {
    throw new Error("Chat hooks must be used within ChatProvider");
  }
  return context;
}

export function useChatComposer() {
  return useChatContext().composer;
}

export function useConversation() {
  return useChatContext().conversation;
}

export function useChatSuggestions() {
  return useChatContext().suggestions;
}

export function useChatModels() {
  return useChatContext().models;
}

export function useChatSettings() {
  return useChatContext().settings;
}

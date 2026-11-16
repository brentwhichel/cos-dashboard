import { useState } from "react";
import { Outlet, useLocation } from "react-router-dom";
import { motion } from "motion/react";
import { InboxIcon } from "lucide-react";
import { Toaster } from "@/components/ui/sonner";
import { HeaderSlotContext } from "@/lib/header-slot";

export function AppShell() {
  const { pathname } = useLocation();
  const [headerSlot, setHeaderSlot] = useState<HTMLDivElement | null>(null);
  return (
    <div className="flex min-h-screen flex-col bg-background">
      <a
        href="#feed"
        className="sr-only focus:not-sr-only focus:fixed focus:left-3 focus:top-3 focus:z-50 focus:rounded-md focus:bg-primary focus:px-3 focus:py-2 focus:text-sm focus:font-medium focus:text-primary-foreground"
      >
        Skip to messages
      </a>
      <header className="sticky top-0 z-30 border-b bg-background/95 backdrop-blur supports-[backdrop-filter]:bg-background/80">
        <div className="mx-auto flex h-14 w-full max-w-5xl items-center gap-3 px-4 sm:px-6">
          <span className="flex size-8 items-center justify-center rounded-md bg-primary text-primary-foreground">
            <InboxIcon className="size-4" aria-hidden="true" />
          </span>
          <span className="min-w-0 truncate text-base font-semibold tracking-tight">
            COS Communications
          </span>
          <div ref={setHeaderSlot} className="ml-auto flex shrink-0 items-center" />
        </div>
      </header>
      <HeaderSlotContext.Provider value={headerSlot}>
      <main className="flex flex-1 flex-col">
        {/* key on pathname → each route re-mounts and replays the entrance.
            Entrance-only (no AnimatePresence/exit): an exit animation around
            <Outlet/> would animate the NEXT route's content, not the leaving one. */}
        <motion.div
          key={pathname}
          initial={{ opacity: 0, y: 6 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.15, ease: "easeOut" }}
          className="flex-1"
        >
          <Outlet />
        </motion.div>
      </main>
      </HeaderSlotContext.Provider>
      <Toaster position="bottom-center" />
    </div>
  );
}

export default AppShell;

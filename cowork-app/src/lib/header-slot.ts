import { createContext } from "react";

/**
 * The right side of the app header. Pages portal page-level status into it
 * (the dashboard's "Last refreshed" stamp) so it stays visible while scrolling.
 */
export const HeaderSlotContext = createContext<HTMLElement | null>(null);

import posthog from "posthog-js";

export function usePostHog() {
    posthog.init("phc_F17ArevVv8vfdkwSQOLztZfVUPexEuRc9NmEaJHWJYB", {
        api_host: "https://us.i.posthog.com",
        defaults: "2025-11-30",
        person_profiles: "always", // or 'identified_only'
    });

    return posthog;
}

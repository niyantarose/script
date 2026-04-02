import { mangaGenreProfile } from "./manga.js";
import { otherGenreProfile } from "./other.js";

const PROFILES = [mangaGenreProfile, otherGenreProfile];

export function resolveGenreProfile(raw) {
  return PROFILES.find(profile => profile.match(raw)) || otherGenreProfile;
}

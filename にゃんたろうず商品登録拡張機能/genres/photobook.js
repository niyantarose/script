export const photobookGenreProfile = {
  id: "photobook",
  label: "photobook",
  requiredSections: [],
  imageBuckets: ["main", "detail"],
  dataFields: [],
  match() {
    return false;
  },
  normalize() {
    return {
      genre: "photobook",
      subGenre: null,
      sections: {},
      warnings: []
    };
  }
};

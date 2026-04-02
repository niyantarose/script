export const artbookGenreProfile = {
  id: "artbook",
  label: "artbook",
  requiredSections: [],
  imageBuckets: ["main", "detail"],
  dataFields: [],
  match() {
    return false;
  },
  normalize() {
    return {
      genre: "artbook",
      subGenre: null,
      sections: {},
      warnings: []
    };
  }
};

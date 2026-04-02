export const novelGenreProfile = {
  id: "novel",
  label: "novel",
  requiredSections: [],
  imageBuckets: ["main", "detail"],
  dataFields: [],
  match() {
    return false;
  },
  normalize() {
    return {
      genre: "novel",
      subGenre: null,
      sections: {},
      warnings: []
    };
  }
};

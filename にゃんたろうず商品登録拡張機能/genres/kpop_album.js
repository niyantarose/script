export const kpop_albumGenreProfile = {
  id: "kpop_album",
  label: "kpop_album",
  requiredSections: [],
  imageBuckets: ["main", "detail"],
  dataFields: [],
  match() {
    return false;
  },
  normalize() {
    return {
      genre: "kpop_album",
      subGenre: null,
      sections: {},
      warnings: []
    };
  }
};

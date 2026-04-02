export const goodsGenreProfile = {
  id: "goods",
  label: "goods",
  requiredSections: [],
  imageBuckets: ["main", "detail"],
  dataFields: [],
  match() {
    return false;
  },
  normalize() {
    return {
      genre: "goods",
      subGenre: null,
      sections: {},
      warnings: []
    };
  }
};

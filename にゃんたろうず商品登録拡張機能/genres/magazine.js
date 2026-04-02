export const magazineGenreProfile = {
  id: "magazine",
  label: "magazine",
  requiredSections: [],
  imageBuckets: ["main", "detail"],
  dataFields: [],
  match() {
    return false;
  },
  normalize() {
    return {
      genre: "magazine",
      subGenre: null,
      sections: {},
      warnings: []
    };
  }
};

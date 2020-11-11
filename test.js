let i = 0;
const uLen = {
  cm: i++,
  m: i++,
  km: i,
};



const convertLen = [
  //cm,m,km
  [1,100, 100_000], //cm
  [0.1, 1, 1_000], //m
  [0.00_001, 0.001, 1], //km
];

const conv = function (i, j) {
  data = [
    [1, 100, 100_000],
    [0.1, 1, 1_000],
    [0.00_001, 0.001, 1],
  ];
  return data[i][j];
};

module.exports = { uLen, convertLen };
//convertLength[km][cm]

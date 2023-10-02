/**
 * 行列を転置する関数
 */
function transposeArray(inputArray) {
  var numRows = inputArray.length;
  var numCols = inputArray[0].length;

  // 新しい2次元配列を作成し、行と列を入れ替える
  var transposedArray = [];
  for (var i = 0; i < numCols; i++) {
    transposedArray[i] = [];
    for (var j = 0; j < numRows; j++) {
      transposedArray[i][j] = inputArray[j][i];
    }
  }

  return transposedArray;
}






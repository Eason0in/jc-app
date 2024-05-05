const SelectRange = Array.from({ length: 10 }, (_, i) => (i + 1) * 5); // 5 10 ... 50

const NumList = Array.from({ length: 9 }, (_, i) => i + 3); // 3 4 5 ... 11

const BeamSelectRange = Array.from({ length: 11 }, (_, i) => i * 5 + 10); // 10 15 20 ... 60

export { SelectRange, NumList, BeamSelectRange };

import { useState } from "react";
import "./index.scss";
import Dropzone from "../Dropzone";
import { BeamSelectRange, NumList } from "../../constants";
import TidyItem from "../TidyItem";

const initACTidyObj = NumList.reduce((obj, num) => {
  return {
    ...obj,
    [`#${num}`]: {
      isNeedTidy: false,
      selectBeamRange: BeamSelectRange[0],
    },
  };
}, {});

function Beam() {
  const [beamMaterialFileA, setBeamMaterialFileA] = useState("");
  const [beamMaterialFile, setBeamMaterialFile] = useState("");
  const [beamTidyFile, setBeamTidyFile] = useState("");
  const [beamTidyFileA, setBeamTidyFileA] = useState("");
  const [beamConstructionFile, setBeamConstructionFile] = useState("");
  const [beamConstructionFileA, setBeamConstructionFileA] = useState("");
  const [beamTidyBySheetNameFile, setBeamTidyBySheetNameFile] = useState("");
  const [beamTidyBySheetNameFileA, setBeamTidyBySheetNameFileA] = useState("");
  const [selectBeamRange, setSelectBeamRange] = useState(BeamSelectRange[0]);
  const [fileName, setFileName] = useState("");
  const [isNeedTidy, setIsNeedTidy] = useState(false);
  const [acTidyObj, setACTidyObj] = useState(initACTidyObj);

  const handleClear = () => {
    setBeamMaterialFile("");
    setBeamMaterialFileA("");
    setBeamTidyFile("");
    setBeamTidyFileA("");
    setBeamConstructionFile("");
    setBeamConstructionFileA("");
    setBeamTidyBySheetNameFile("");
    setBeamTidyBySheetNameFileA("");
    setFileName("");
  };

  const handleInputChange = (e) => {
    const [file] = e;
    setFileName(() => file.name);
    const data = {
      filePath: file.path,
      range: selectBeamRange,
      isNeedTidy,
      acTidyObj,
    };
    window.electronAPI.beamReadFile(data);
  };
  const handleSelectChange = (e) => {
    setSelectBeamRange(e.target.value);
  };

  const handleNeedTidy = () => {
    setIsNeedTidy(!isNeedTidy);
  };

  const handleACNeedTidy = ({ num, type, value }) => {
    setACTidyObj((last) => {
      return {
        ...last,
        [`#${num}`]: {
          ...last[`#${num}`],
          [type]: value,
        },
      };
    });
  };

  window.electronAPI.sendBeamMaterialFile((event, content) => {
    const blobData = new Blob([content], {
      type: "application/vnd.ms-excel;charset=utf-8;",
    });

    setBeamMaterialFile(URL.createObjectURL(blobData));
    setBeamMaterialFileA("料單.xlsx");
  });

  window.electronAPI.sendBeamTidyFile((event, content) => {
    const blobData = new Blob([content], {
      type: "application/vnd.ms-excel;charset=utf-8;",
    });

    setBeamTidyFile(URL.createObjectURL(blobData));
    setBeamTidyFileA("歸整.xlsx");
  });

  window.electronAPI.sendBeamTidyBySheetNameFile((event, content) => {
    const blobData = new Blob([content], {
      type: "application/vnd.ms-excel;charset=utf-8;",
    });

    setBeamTidyBySheetNameFile(URL.createObjectURL(blobData));
    setBeamTidyBySheetNameFileA("料單下方歸整統計表.xlsx");
  });
  window.electronAPI.sendBeamConstructionFile((event, content) => {
    const blobData = new Blob([content], {
      type: "application/vnd.ms-excel;charset=utf-8;",
    });

    setBeamConstructionFile(URL.createObjectURL(blobData));
    setBeamConstructionFileA("歸整後施工圖.xlsx");
  });

  return (
    <section id="beam">
      <div className="needTidy">
        <label htmlFor="isNeedTidy">是否需要歸整(CC例外)</label>
        <input
          id="isNeedTidy"
          type="checkbox"
          value={isNeedTidy}
          onClick={handleNeedTidy}
        />
      </div>

      <div className="range">
        <label htmlFor="beamRange">歸整間距(CC例外)：</label>
        <select
          name="beamRange"
          id="beamRange"
          value={selectBeamRange}
          onChange={handleSelectChange}
        >
          {BeamSelectRange.map((range) => (
            <option key={range} value={range}>
              {range}
            </option>
          ))}
        </select>
      </div>

      <button onClick={handleClear}>清除檔案</button>

      <div className="acTidy">
        <label htmlFor="isACNeedTidy">AC 規整</label>
        <div className="acNeedTidy">
          {NumList.map((item, index) => (
            <TidyItem
              key={index}
              num={item}
              {...acTidyObj[`#${item}`]}
              handleNeedTidy={handleACNeedTidy}
            />
          ))}
        </div>
      </div>

      <Dropzone
        classArr="dropZone"
        fileName={fileName}
        accept=".xlsx"
        handleInputChange={handleInputChange}
      />

      <label className="fileName">檔案:</label>
      <ul>
        <li>
          <a href={beamTidyFile} download="歸整.xlsx">
            {beamTidyFileA}
          </a>
        </li>
        <li>
          <a href={beamMaterialFile} download="料單.xlsx">
            {beamMaterialFileA}
          </a>
        </li>
        <li>
          <a href={beamTidyBySheetNameFile} download="料單下方歸整統計表.xlsx">
            {beamTidyBySheetNameFileA}
          </a>
        </li>
        <li>
          <a href={beamConstructionFile} download="歸整後施工圖.xlsx">
            {beamConstructionFileA}
          </a>
        </li>
      </ul>
    </section>
  );
}

export default Beam;

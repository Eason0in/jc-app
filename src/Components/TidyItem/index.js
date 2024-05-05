import "./index.scss";
import { BeamSelectRange } from "../../constants";

function TidyItem({ num, isNeedTidy, handleNeedTidy, selectBeamRange }) {
  return (
    <div className="tidyItem">
      <label htmlFor={`isACNeedTidy${num}`}>#{num}</label>
      <input
        id={`isACNeedTidy${num}`}
        type="checkbox"
        value={isNeedTidy}
        onClick={() =>
          handleNeedTidy({ num, type: "isNeedTidy", value: !isNeedTidy })
        }
      />

      <select
        name="beamRange"
        value={selectBeamRange}
        onChange={(e) =>
          handleNeedTidy({
            num,
            type: "selectBeamRange",
            value: Number(e.target.value),
          })
        }
      >
        {BeamSelectRange.map((range) => (
          <option key={range} value={range}>
            {range}
          </option>
        ))}
      </select>
    </div>
  );
}

export default TidyItem;

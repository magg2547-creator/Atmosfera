export function calcAqiFromPm25(pm25) {
  const breakpoints = [
    [0, 12, 0, 50],
    [12.1, 35.4, 51, 100],
    [35.5, 55.4, 101, 150],
    [55.5, 150.4, 151, 200],
    [150.5, 250.4, 201, 300],
    [250.5, 350.4, 301, 400],
    [350.5, 500.4, 401, 500],
  ];

  for (const [low, high, aqiLow, aqiHigh] of breakpoints) {
    if (pm25 >= low && pm25 <= high) {
      return Math.round(((aqiHigh - aqiLow) / (high - low)) * (pm25 - low) + aqiLow);
    }
  }

  return 500;
}

export function calcAqiFromPm10(pm10) {
  const breakpoints = [
    [0, 54, 0, 50],
    [55, 154, 51, 100],
    [155, 254, 101, 150],
    [255, 354, 151, 200],
    [355, 424, 201, 300],
    [425, 504, 301, 400],
    [505, 604, 401, 500],
  ];

  for (const [low, high, aqiLow, aqiHigh] of breakpoints) {
    if (pm10 >= low && pm10 <= high) {
      return Math.round(((aqiHigh - aqiLow) / (high - low)) * (pm10 - low) + aqiLow);
    }
  }

  return 500;
}

export function calcAQI(pm25, pm10) {
  return Math.max(
    calcAqiFromPm25(pm25),
    calcAqiFromPm10(pm10)
  );
}

export function getAqiMeta(aqi) {
  if (aqi <= 50) return { label: 'Good', dot: '#a8e6c8' };
  if (aqi <= 100) return { label: 'Moderate', dot: '#fde68a' };
  if (aqi <= 150) return { label: 'Unhealthy for Sensitive Groups', dot: '#fca5a5' };
  if (aqi <= 200) return { label: 'Unhealthy', dot: '#f87171' };
  if (aqi <= 300) return { label: 'Very Unhealthy', dot: '#c084fc' };
  return { label: 'Hazardous', dot: '#9f1239' };
}

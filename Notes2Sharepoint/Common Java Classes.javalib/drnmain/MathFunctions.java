package drnmain;

public class MathFunctions {

	public static double getDistance(double lat1, String lat1S, double lon1,
			String lon1S, double lat2, String lat2S, double lon2, String lon2S) {
		double a = 6378137;
		double b = 6356752.3142;
		double f = 1 / 298.257223563;
		double L;
		double U1, U2, sinU1, sinU2, lambda, lambdaP, iterLimit;
		double sinLambda, sinSigma, cosSigma, sigma, sinAlpha, cosSqAlpha, cos2SigmaM;
		double cosLambda;
		double C;
		double uSq, A, B, deltaSigma, s;
		double cosU1, cosU2;
		double NM;
		Loger outLog = new Loger(false);
		

		outLog.s("Start of getDistance");
		outLog.s(lat1 + lat1S + " " + lon1 + lon1S);
		outLog.s(lat2 + lat2S + " " + lon2 + lon2S);
		s = 0;
		if (lat1S.equals("S")) {
			lat1 = -lat1;

		}
		if (lat2S.equals("S")) {
			lat2 = -lat2;
		}
		if (lon1S.equals("W")) {
			lon1 = -lon1;
		}
		if (lon2S.equals("W")) {
			lon2 = -lon2;
		}

		L = Math.toRadians(lon2 - lon1);
		U1 = Math.atan((1 - f) * Math.tan(Math.toRadians(lat1)));
		U2 = Math.atan((1 - f) * Math.tan(Math.toRadians(lat2)));
		sinU1 = Math.sin(U1);
		cosU1 = Math.cos(U1);
		sinU2 = Math.sin(U2);
		cosU2 = Math.cos(U2);

		lambda = L;
		iterLimit = 100;
		do {
			sinLambda = Math.sin(lambda);
			cosLambda = Math.cos(lambda);
			sinSigma = Math.sqrt((cosU2 * sinLambda) * (cosU2 * sinLambda)
					+ (cosU1 * sinU2 - sinU1 * cosU2 * cosLambda)
					* (cosU1 * sinU2 - sinU1 * cosU2 * cosLambda));
			if (sinSigma == 0) {
				outLog.s("sinSigma = 0");
				return 0;
			}
			cosSigma = sinU1 * sinU2 + cosU1 * cosU2 * cosLambda;
			sigma = Math.atan2(sinSigma, cosSigma);
			sinAlpha = cosU1 * cosU2 * sinLambda / sinSigma;
			cosSqAlpha = 1 - sinAlpha * sinAlpha;
			// s("cosSigma = " + cosSigma);
			// s("sinU1 " + sinU1 + " sinU2 " + sinU2);
			// s("cosSqAlpha " + cosSqAlpha);
			cos2SigmaM = cosSigma - 2 * sinU1 * sinU2 / cosSqAlpha;

			if (Double.isNaN(cos2SigmaM)) {
				outLog.s("cos2SigmaM is not a number " + cos2SigmaM);
				return -1;
			}

			C = f / 16 * cosSqAlpha * (4 + f * (4 - 3 * cosSqAlpha));
			lambdaP = lambda;
			lambda = L
					+ (1 - C)
					* f
					* sinAlpha
					* (sigma + C
							* sinSigma
							* (cos2SigmaM + C * cosSigma
									* (-1 + 2 * cos2SigmaM * cos2SigmaM)));

		} while (Math.abs(lambda - lambdaP) > 1e-12 && (--iterLimit > 0));
		if (iterLimit == 0) {
			outLog.s("failed to converge iterLimit = " + iterLimit);
			return -1;
		}
		uSq = cosSqAlpha * (a * a - b * b) / (b * b);
		A = 1 + uSq / 16384 * (4096 + uSq * (-768 + uSq * (320 - 175 * uSq)));
		B = uSq / 1024 * (256 + uSq * (-128 + uSq * (74 - 47 * uSq)));
		// s("ssig " + sinSigma + " c2sigM " + cos2SigmaM);
		// s("B " + B);
		deltaSigma = B
				* sinSigma
				* (cos2SigmaM + B
						/ 4
						* (cosSigma * (-1 + 2 * cos2SigmaM * cos2SigmaM) - B
								/ 6 * cos2SigmaM
								* (-3 + 4 * sinSigma * sinSigma)
								* (-3 + 4 * cos2SigmaM * cos2SigmaM)));
		// s("A = " + A + " b = " + b);
		// s("sigma = " + sigma);
		// s("deltaSigma = " + deltaSigma);
		s = b * A * (sigma - deltaSigma);
		s = s / 1000;
		outLog.s("KM = " + s);

		NM = s * .53996;
		outLog.s("NM = " + NM);

		outLog.s("getDistance finished");
		return NM;
	}

	/**
	 * Round a double to two decimal places using normal rounding rules where
	 * over .5 rounds up Java rounding causes rounding to 0 this is required
	 * frequently because even simple math in java frequently results in values
	 * small fraction in error.
	 * 
	 * Copyright Don Newman Author: Don Newman Creation date: (01/11/2001
	 * 8:32:16 AM)
	 * 
	 * @return double
	 * @param input
	 *            double
	 */

	public static double drnRoundMoney(double valIn) {

		double result;

		result = valIn * 100.0;

		result = Math.round(result);
		result = result / 100.0;

		return result;
	}
	
	/**
	 * Round the first double to the second double number of decimal places.
	 * 
	 * @param valIn
	 * @param decPlaces
	 * @return Creation date: 2-Feb-10 3:33:33 PM author: Don Newman copyright
	 *         <b><i>Don Newman</b></i>
	 */
	public static double roundNumber(double valIn, double decPlaces) {

		double result;

		double multiplier = Math.pow(10, decPlaces);

		result = valIn * multiplier;

		result = Math.round(result);
		result = result / multiplier;

		return result;
	}

	/**
	 * 
	 * @param degrees
	 * @param minutes
	 * @return Creation date: 5-May-09 9:10:22 AM author: Don Newman copyright
	 *         <b><i>Don Newman</b></i>
	 */
	public static double getDecimalDegrees(double degrees, double minutes) {
		return degrees + (minutes / 60);
	}

}

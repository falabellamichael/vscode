/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { localize } from '../../../nls.js';
import { AssistantWeatherDay, AssistantWeatherForecast, AssistantWeatherLocation, normalizeAssistantWeatherLocation } from '../common/native.js';

type AssistantWeatherFetcher = (url: string, options: RequestInit) => Promise<Response>;

const forecastEndpoint = 'https://api.data.pelmorex.com/weather/LongTermForecasts/';
const forecastPage = 'https://www.theweathernetwork.com/en/city';

interface PelmorexLocation {
	readonly Name?: string;
	readonly SubDivision?: string;
	readonly SubDivisionCode?: string;
}

function forecastSourceUrl(country: string, location: PelmorexLocation | undefined): string {
	const slug = (value: string | undefined) => typeof value === 'string'
		? value.normalize('NFD').replace(/\p{M}/gu, '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '')
		: '';
	const city = slug(location?.Name);
	const subdivision = slug(location?.SubDivision);
	if ((country === 'CA' || country === 'US') && city && subdivision) {
		return forecastPage + '/' + country.toLowerCase() + '/' + subdivision + '/' + city + '/7-days';
	}
	return forecastPage;
}

function numberField(value: unknown): number | undefined {
	if (typeof value !== 'number' && (typeof value !== 'string' || !/^-?\d+(?:\.\d+)?$/.test(value.trim()))) { return undefined; }
	const number = Number(value);
	return Number.isFinite(number) ? number : undefined;
}

export async function fetchAssistantWeatherForecast(day: AssistantWeatherDay, location: AssistantWeatherLocation, key: string, fetcher: AssistantWeatherFetcher, allowPublicFallback = false): Promise<AssistantWeatherForecast> {
	if (day !== 'today' && day !== 'tomorrow') { throw new Error(localize('assistantWeather.invalidDay', "Choose today or tomorrow for the forecast.")); }
	const normalizedLocation = normalizeAssistantWeatherLocation(location);
	if (!normalizedLocation) {
		throw new Error(localize('assistantWeather.invalidLocation', "Enter a city, a two-letter country code, and a province, state, or region code if needed."));
	}
	if (typeof key === 'string' && /^[\x21-\x7e]{1,512}$/.test(key)) {
		try {
			return await fetchPelmorexForecast(day, normalizedLocation, key, fetcher);
		} catch (error) {
			// Key verification must see the provider's own failure; briefings may
			// fall back to the public Open-Meteo service when it is unavailable.
			if (!allowPublicFallback) { throw error; }
		}
	} else if (!allowPublicFallback) {
		throw new Error('Enter a valid Pelmorex API key.');
	}
	return fetchOpenMeteoForecast(day, normalizedLocation, fetcher);
}

async function fetchPelmorexForecast(day: AssistantWeatherDay, normalizedLocation: AssistantWeatherLocation, key: string, fetcher: AssistantWeatherFetcher): Promise<AssistantWeatherForecast> {
	const { city, country, subdivision } = normalizedLocation;
	const locationLabel = [city, subdivision, country].filter(Boolean).join(', ');
	const path = [country, subdivision, encodeURIComponent(city)].filter(Boolean).join('/');
	let response: Response;
	try {
		response = await fetcher(forecastEndpoint + path + '?locale=en-CA', {
			headers: { 'Accept': 'application/json', 'x-api-key': key },
			credentials: 'omit',
			redirect: 'error',
			signal: AbortSignal.timeout(8_000),
		});
	} catch {
		throw new Error('The Weather Network forecast request failed or timed out.');
	}
	if (!response.ok) {
		void response.body?.cancel().catch(() => {});
		if (response.status === 401 || response.status === 403) { throw new Error('Pelmorex rejected the API key or it lacks access to Long Term Forecasts. Check the key with /weather.'); }
		if (response.status === 404) { throw new Error(localize('assistantWeather.locationNotFound', "The Weather Network could not find {0}. Check the city and its province, state, or region.", locationLabel)); }
		throw new Error(`The Weather Network forecast is unavailable (HTTP ${response.status}).`);
	}
	const data: { LongTermPeriod?: (Record<string, unknown> | null)[]; Location?: PelmorexLocation } | null = await response.json().catch(() => null);
	const periods = Array.isArray(data?.LongTermPeriod) ? data.LongTermPeriod : [];
	// Pelmorex defines period 0 as today in the forecast location.
	const dayIndex = day === 'tomorrow' ? 1 : 0;
	const indexedPeriod = periods[dayIndex];
	const period = periods.find(item => numberField(item?.Period) === dayIndex)
		?? (indexedPeriod?.Period === undefined || indexedPeriod.Period === null ? indexedPeriod : undefined);
	const date = typeof period?.TimestampLocal === 'string' ? period.TimestampLocal.slice(0, 10) : '';
	if (!period || !/^\d{4}-\d{2}-\d{2}$/.test(date)) {
		throw new Error(localize('assistantWeather.missingForecast', "The Weather Network did not return the requested forecast day for {0}.", locationLabel));
	}
	const highC = numberField(period.TemperatureCMax);
	const lowC = numberField(period.TemperatureCMin);
	const probability = numberField(period.POPPercentDay);
	const precipitationChancePercent = probability !== undefined && probability >= 0 && probability <= 100 ? probability : undefined;
	const conditions = typeof period.FxConditionDay === 'string' ? period.FxConditionDay.trim().slice(0, 160) || undefined : undefined;
	if (highC === undefined && lowC === undefined && precipitationChancePercent === undefined && !conditions) {
		throw new Error(localize('assistantWeather.emptyForecast', "The Weather Network returned an empty forecast for {0} on {1}.", locationLabel, date));
	}
	const providerLocation = data?.Location;
	const name = typeof providerLocation?.Name === 'string' ? providerLocation.Name.trim() : '';
	const region = typeof providerLocation?.SubDivisionCode === 'string' ? providerLocation.SubDivisionCode.trim() : '';
	return {
		source: 'The Weather Network',
		sourceUrl: forecastSourceUrl(country, providerLocation),
		location: [name || city, region || subdivision, country].filter(Boolean).join(', '),
		date, highC, lowC, precipitationChancePercent, conditions,
	};
}

const openMeteoGeocodingEndpoint = 'https://geocoding-api.open-meteo.com/v1/search';
const openMeteoForecastEndpoint = 'https://api.open-meteo.com/v1/forecast';

interface OpenMeteoGeoResult {
	readonly name?: string;
	readonly country_code?: string;
	readonly admin1?: string;
	readonly latitude?: number;
	readonly longitude?: number;
}

// Open-Meteo geocoding reports region names rather than codes; map the CA/US
// subdivision codes accepted by normalizeAssistantWeatherLocation so a
// province or state still disambiguates cities with the same name.
const openMeteoRegionNames: Record<string, string> = {
	'CA:AB': 'Alberta', 'CA:BC': 'British Columbia', 'CA:MB': 'Manitoba', 'CA:NB': 'New Brunswick',
	'CA:NL': 'Newfoundland and Labrador', 'CA:NS': 'Nova Scotia', 'CA:NT': 'Northwest Territories',
	'CA:NU': 'Nunavut', 'CA:ON': 'Ontario', 'CA:PE': 'Prince Edward Island', 'CA:QC': 'Quebec',
	'CA:SK': 'Saskatchewan', 'CA:YT': 'Yukon',
	'US:AL': 'Alabama', 'US:AK': 'Alaska', 'US:AZ': 'Arizona', 'US:AR': 'Arkansas', 'US:CA': 'California',
	'US:CO': 'Colorado', 'US:CT': 'Connecticut', 'US:DE': 'Delaware', 'US:DC': 'District of Columbia',
	'US:FL': 'Florida', 'US:GA': 'Georgia', 'US:HI': 'Hawaii', 'US:ID': 'Idaho', 'US:IL': 'Illinois',
	'US:IN': 'Indiana', 'US:IA': 'Iowa', 'US:KS': 'Kansas', 'US:KY': 'Kentucky', 'US:LA': 'Louisiana',
	'US:ME': 'Maine', 'US:MD': 'Maryland', 'US:MA': 'Massachusetts', 'US:MI': 'Michigan', 'US:MN': 'Minnesota',
	'US:MS': 'Mississippi', 'US:MO': 'Missouri', 'US:MT': 'Montana', 'US:NE': 'Nebraska', 'US:NV': 'Nevada',
	'US:NH': 'New Hampshire', 'US:NJ': 'New Jersey', 'US:NM': 'New Mexico', 'US:NY': 'New York',
	'US:NC': 'North Carolina', 'US:ND': 'North Dakota', 'US:OH': 'Ohio', 'US:OK': 'Oklahoma', 'US:OR': 'Oregon',
	'US:PA': 'Pennsylvania', 'US:RI': 'Rhode Island', 'US:SC': 'South Carolina', 'US:SD': 'South Dakota',
	'US:TN': 'Tennessee', 'US:TX': 'Texas', 'US:UT': 'Utah', 'US:VT': 'Vermont', 'US:VA': 'Virginia',
	'US:WA': 'Washington', 'US:WV': 'West Virginia', 'US:WI': 'Wisconsin', 'US:WY': 'Wyoming',
};

const openMeteoConditions: Record<number, string> = {
	0: 'Clear sky', 1: 'Mainly clear', 2: 'Partly cloudy', 3: 'Overcast',
	45: 'Fog', 48: 'Depositing rime fog',
	51: 'Light drizzle', 53: 'Drizzle', 55: 'Dense drizzle',
	56: 'Light freezing drizzle', 57: 'Freezing drizzle',
	61: 'Slight rain', 63: 'Rain', 65: 'Heavy rain',
	66: 'Light freezing rain', 67: 'Freezing rain',
	71: 'Slight snow', 73: 'Snow', 75: 'Heavy snow', 77: 'Snow grains',
	80: 'Slight rain showers', 81: 'Rain showers', 82: 'Violent rain showers',
	85: 'Slight snow showers', 86: 'Snow showers',
	95: 'Thunderstorm', 96: 'Thunderstorm with slight hail', 99: 'Thunderstorm with hail',
};

function openMeteoRegionMatches(result: OpenMeteoGeoResult, country: string, subdivision: string | undefined): boolean {
	if (!subdivision) { return true; }
	const admin1 = typeof result.admin1 === 'string' ? result.admin1.trim().toLowerCase() : '';
	if (!admin1) { return false; }
	const name = openMeteoRegionNames[`${country}:${subdivision}`]?.toLowerCase();
	return admin1 === subdivision.toLowerCase() || (name !== undefined && admin1 === name);
}

async function fetchOpenMeteoForecast(day: AssistantWeatherDay, normalizedLocation: AssistantWeatherLocation, fetcher: AssistantWeatherFetcher): Promise<AssistantWeatherForecast> {
	const { city, country, subdivision } = normalizedLocation;
	const locationLabel = [city, subdivision, country].filter(Boolean).join(', ');
	const requestJson = async (url: string, failureLabel: string): Promise<unknown> => {
		let response: Response;
		try {
			response = await fetcher(url, {
				headers: { 'Accept': 'application/json' },
				credentials: 'omit',
				redirect: 'error',
				signal: AbortSignal.timeout(8_000),
			});
		} catch {
			throw new Error(`The Open-Meteo ${failureLabel} request failed or timed out.`);
		}
		if (!response.ok) {
			void response.body?.cancel().catch(() => {});
			throw new Error(`The Open-Meteo ${failureLabel} is unavailable (HTTP ${response.status}).`);
		}
		return response.json().catch(() => null);
	};
	const geocoding = await requestJson(`${openMeteoGeocodingEndpoint}?name=${encodeURIComponent(city)}&count=10&language=en&format=json`, 'location search') as { results?: (OpenMeteoGeoResult | null)[] } | null;
	const candidates = (Array.isArray(geocoding?.results) ? geocoding.results : []).filter((result): result is OpenMeteoGeoResult & { name: string } =>
		!!result && typeof result.name === 'string' && !!result.name.trim()
		&& typeof result.country_code === 'string' && result.country_code.toUpperCase() === country
		&& typeof result.latitude === 'number' && Number.isFinite(result.latitude)
		&& typeof result.longitude === 'number' && Number.isFinite(result.longitude));
	// Prefer the requested province/state; otherwise keep the provider's top-ranked
	// city in the country so a region name it does not recognize still gets a forecast.
	const place = candidates.find(result => openMeteoRegionMatches(result, country, subdivision)) ?? candidates[0];
	if (!place) {
		throw new Error(localize('assistantWeather.openMeteoLocationNotFound', "Open-Meteo could not find {0}. Check the city and its province, state, or region.", locationLabel));
	}
	const forecast = await requestJson(`${openMeteoForecastEndpoint}?latitude=${place.latitude}&longitude=${place.longitude}&daily=temperature_2m_max,temperature_2m_min,precipitation_probability_max,weather_code&timezone=auto&forecast_days=2`, 'forecast') as { daily?: Record<string, unknown> } | null;
	const daily = forecast?.daily;
	const dayIndex = day === 'tomorrow' ? 1 : 0;
	const dailyField = (key: string): unknown => Array.isArray(daily?.[key]) ? (daily[key] as unknown[])[dayIndex] : undefined;
	// timezone=auto reports each day in the forecast location, matching Pelmorex period 0 = today.
	const date = typeof dailyField('time') === 'string' ? (dailyField('time') as string).slice(0, 10) : '';
	if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
		throw new Error(localize('assistantWeather.openMeteoMissingForecast', "Open-Meteo did not return the requested forecast day for {0}.", locationLabel));
	}
	const highC = numberField(dailyField('temperature_2m_max'));
	const lowC = numberField(dailyField('temperature_2m_min'));
	const probability = numberField(dailyField('precipitation_probability_max'));
	const precipitationChancePercent = probability !== undefined && probability >= 0 && probability <= 100 ? probability : undefined;
	const weatherCode = numberField(dailyField('weather_code'));
	const conditions = weatherCode !== undefined ? openMeteoConditions[weatherCode] : undefined;
	if (highC === undefined && lowC === undefined && precipitationChancePercent === undefined && !conditions) {
		throw new Error(localize('assistantWeather.openMeteoEmptyForecast', "Open-Meteo returned an empty forecast for {0} on {1}.", locationLabel, date));
	}
	const region = typeof place.admin1 === 'string' ? place.admin1.trim() : '';
	return {
		source: 'Open-Meteo',
		sourceUrl: `https://open-meteo.com/en/docs#latitude=${place.latitude}&longitude=${place.longitude}`,
		location: [place.name.trim(), region || subdivision, country].filter(Boolean).join(', '),
		date, highC, lowC, precipitationChancePercent, conditions,
	};
}

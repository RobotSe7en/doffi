/** Word theme modifiers change HSL luminance; tint takes precedence. */
export function wordColorLuminance(value: string, tint?: string, shade?: string): string {
    const modifier = tint ?? shade;
    if (modifier === undefined)
        return value;
    // Word applies tint in preference to shade and changes HSL luminance.
    const rgb = [0, 2, 4].map((start) => parseInt(value.slice(start, start + 2), 16) / 255);
    const [red, green, blue] = rgb as [
        number,
        number,
        number
    ];
    const maximum = Math.max(...rgb), minimum = Math.min(...rgb), delta = maximum - minimum;
    const lightness = (maximum + minimum) / 2;
    const saturation = delta === 0 ? 0 : delta / (1 - Math.abs(2 * lightness - 1));
    const hue = delta === 0
        ? 0
        : maximum === red
            ? ((green - blue) / delta + 6) % 6
            : maximum === green
                ? (blue - red) / delta + 2
                : (red - green) / delta + 4;
    const factor = parseInt(modifier, 16) / 255;
    const resultLightness = tint !== undefined ? 1 - (1 - lightness) * factor : lightness * factor;
    const chroma = (1 - Math.abs(2 * resultLightness - 1)) * saturation;
    const x = chroma * (1 - Math.abs((hue % 2) - 1)), offset = resultLightness - chroma / 2;
    const channels = hue < 1
        ? [chroma, x, 0]
        : hue < 2
            ? [x, chroma, 0]
            : hue < 3
                ? [0, chroma, x]
                : hue < 4
                    ? [0, x, chroma]
                    : hue < 5
                        ? [x, 0, chroma]
                        : [chroma, 0, x];
    return channels
        .map((channel) => Math.round((channel + offset) * 255)
        .toString(16)
        .padStart(2, "0"))
        .join("");
}

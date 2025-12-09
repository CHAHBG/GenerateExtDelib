/** @type {import('tailwindcss').Config} */
export default {
    content: [
        "./index.html",
        "./src/**/*.{js,ts,jsx,tsx}",
    ],
    theme: {
        extend: {
            colors: {
                primary: '#2563eb', // Royal Blue
                secondary: '#0f172a', // Slate 900
                accent: '#f59e0b', // Amber
            }
        },
    },
    plugins: [],
}

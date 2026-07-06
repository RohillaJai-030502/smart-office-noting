// ============================================================
// transliterate.js — Offline English-to-Hindi Phonetic Transliterator
// ============================================================

const CONSONANTS = {
    "ksh": "क्ष", "shh": "ष", "gya": "ज्ञ", "chr": "क्र",
    "kh": "ख", "gh": "घ", "ch": "च", "chh": "छ", "jh": "झ",
    "th": "थ", "dh": "ध", "ph": "फ", "bh": "भ", "sh": "श",
    "k": "क", "g": "ग", "j": "ज", "t": "त", "d": "द", "n": "न",
    "p": "प", "b": "ब", "m": "म", "y": "य", "r": "र", "l": "ल",
    "v": "व", "w": "व", "s": "स", "h": "ह", "f": "फ़", "z": "ज़"
};

const VOWELS = {
    "aa": "ा", "ai": "ै", "au": "ौ", "ee": "ी", "oo": "ू",
    "a": "", "i": "ि", "u": "ु", "e": "े", "o": "ो"
};

const INITIAL_VOWELS = {
    "aa": "आ", "ai": "ऐ", "au": "औ", "ee": "ई", "oo": "ऊ",
    "a": "अ", "i": "इ", "u": "उ", "e": "ए", "o": "ओ"
};

// Dictionary of common office terms and names for perfect offline accuracy
const COMMON_DICTIONARY = {
    "rakesh": "राकेश",
    "kumar": "कुमार",
    "amit": "अमित",
    "sharma": "शर्मा",
    "neha": "नेहा",
    "gupta": "गुप्ता",
    "singh": "सिंह",
    "dr": "डॉ.",
    "dr.": "डॉ.",
    "anil": "अनिल",
    "khurana": "खुराना",
    "hrdd": "मा.सं.वि.वि.",
    "director": "निदेशक",
    "scientist": "वैज्ञानिक",
    "trainee": "प्रशिक्षु",
    "paid": "पेड",
    "internship": "इंटर्नशिप",
    "scheme": "स्कीम",
    "stipend": "स्टाइपेंड",
    "payment": "भुगतान",
    "visit": "विजिट",
    "faculty": "फैकल्टी",
    "mentor": "मेंटर",
    "nomination": "नामांकन",
    "cancellation": "निरस्तीकरण",
    "date": "तिथि",
    "amendment": "संशोधन",
    "revision": "संशोधन",
    "information": "सूचना",
    "project": "परियोजना",
    "coordinator": "समन्वयक",
    "provisional": "अनंतिम",
    "certificate": "प्रमाण पत्र"
};

function phoneticTransliterateWord(word) {
    if (!word) return "";
    let lower = word.toLowerCase();
    
    // Check dictionary first
    if (COMMON_DICTIONARY[lower]) {
        return COMMON_DICTIONARY[lower];
    }
    
    let result = "";
    let i = 0;
    let lastWasConsonant = false;
    
    while (i < lower.length) {
        let matchFound = false;
        
        // 1. Check Vowels (longest match first)
        let vowelKeys = ["aa", "ai", "au", "ee", "oo", "a", "i", "u", "e", "o"];
        for (let vk of vowelKeys) {
            if (lower.startsWith(vk, i)) {
                let isInitial = (i === 0 || !lastWasConsonant);
                // Smart rule: final "a" in Indian names is usually pronounced "aa" (e.g. Sharma -> शर्मा, Neha -> नेहा)
                if (vk === "a" && i === lower.length - 1 && lower.length > 2) {
                    result += "ा";
                } else {
                    result += isInitial ? (INITIAL_VOWELS[vk] || "") : (VOWELS[vk] || "");
                }
                i += vk.length;
                lastWasConsonant = false;
                matchFound = true;
                break;
            }
        }
        if (matchFound) continue;
        
        // 2. Check Consonants (longest match first)
        let consKeys = ["ksh", "shh", "gya", "chr", "kh", "gh", "ch", "chh", "jh", "th", "dh", "ph", "bh", "sh", "k", "g", "j", "t", "d", "n", "p", "b", "m", "y", "r", "l", "v", "w", "s", "h", "f", "z"];
        for (let ck of consKeys) {
            if (lower.startsWith(ck, i)) {
                let char = CONSONANTS[ck];
                
                // Smart rule: if this consonant is followed by another consonant (not a vowel), it's a half-consonant (halant)
                let nextIsVowel = false;
                let nextIdx = i + ck.length;
                if (nextIdx < lower.length) {
                    let nextChar = lower[nextIdx];
                    if (["a", "e", "i", "o", "u"].includes(nextChar)) {
                        nextIsVowel = true;
                    }
                }
                
                // Do not add halant if it's the last character
                if (!nextIsVowel && nextIdx < lower.length) {
                    // Avoid halant for "r" followed by consonant in some contexts, but standard halant works
                    result += char + "्";
                } else {
                    result += char;
                }
                
                i += ck.length;
                lastWasConsonant = true;
                matchFound = true;
                break;
            }
        }
        
        if (!matchFound) {
            // Keep punctuation/numbers as is
            result += lower[i];
            i++;
            lastWasConsonant = false;
        }
    }
    
    return result;
}

function transliterateEnglishToHindi(text) {
    if (!text) return "";
    let parts = text.split(/(\s+)/);
    let output = [];
    for (let part of parts) {
        if (/^\s+$/.test(part)) {
            output.push(part);
        } else {
            output.push(phoneticTransliterateWord(part));
        }
    }
    return output.join("");
}

// Global hook for binding inputs
function setupTransliteration(sourceId, targetId, toggleId) {
    const sourceEl = document.getElementById(sourceId);
    const targetEl = document.getElementById(targetId);
    const toggleEl = document.getElementById(toggleId);
    
    if (!sourceEl || !targetEl) return;
    
    sourceEl.addEventListener("input", () => {
        if (toggleEl && !toggleEl.checked) return;
        targetEl.value = transliterateEnglishToHindi(sourceEl.value);
        // Trigger input event on target to update live preview
        targetEl.dispatchEvent(new Event("input"));
    });
}

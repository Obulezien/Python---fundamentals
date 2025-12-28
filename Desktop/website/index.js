
/*function rollDice(){
    const inputDice = document.getElementById("inputDice").value;
    const diceResults = document.getElementById("diceResults");
    const diceImages = document.getElementById("diceImages");
    const diceArray = [];
    const imageArray = [];
        
    for(let i = 0; i < inputDice; i++){
        const value = Math.floor(Math.random()*6)+1;   
        diceArray.push(value);      
        imageArray.push(`<img style="border-radius:15px; padding:2px;" src="dice_images/${value}.png" alt="Dice: ${value}">`);   
    }
    diceResults.textContent = `Dice: ${diceArray.join(", ")}`;
    
    diceImages.innerHTML = imageArray.join(" ");
}*/
function generatePassword(passlength, 
                            includeLowerCaseChars,
                            includeUpperCaseChars,
                            includeNumbers,
                            includeSymbols){
            const lowerCaseChars = "abcdefghijklmnopqrstuvwxyz";
            const upperCaseChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            const numbers = "123456789";
            const symbols = "!*%$£_-()+?/\|&^";

            let allowedChars = "";
            let generatedPassword = "";

            allowedChars += includeLowerCaseChars ? lowerCaseChars : "";
            allowedChars += includeUpperCaseChars ? upperCaseChars : "";
            allowedChars += includeNumbers ? numbers : "";
            allowedChars += includeSymbols ? symbols : "";

            if(passlength <= 0){
                return `passlength must be greater 0`;
            }

            if(allowedChars === 0){
                return "You need to enable at least 1 set of characters";
            }

            for(let i = 0; i < passlength; i++){
                charIndex = Math.floor(Math.random()*allowedChars.length);
                generatedPassword += allowedChars[charIndex];
            }

            return generatedPassword;
    }

    let generatedPassword = generatePassword(12, true, true, true, true);

    console.log(`Password: ${generatedPassword}`);
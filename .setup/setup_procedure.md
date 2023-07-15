# Setup commands

1. Create `.venv`

    ```powershell
    python -m venv .venv
    ```

2. Activate `.venv`

    For bash:

    ```bash
    .venv\Scripts\activate
    ```

    For powershell:

    ```powershell
    .venv\Scripts\Activate.ps1
    ```

3. Install [nodeenv](https://github.com/ekalinin/nodeenv) for Node.js virtual environment

    ```powershell
    pip install nodeenv
    ```

4. Create `.env` Node.js virtual environment

    ```powershell
    nodeenv .env
    ```

5. Activate Node.js virtual environment
    For bash:

    ```bash
    .env\Scripts\activate
    ```

    For powershell:

    ```powershell
    .env\Scripts\Activate.ps1
    ```

6. Install Node.js dependencies

    ```powershell
    npm install @google/clasp -g
    npm i -S @types/google-apps-script
    ```

7. Logs in and authorize management of your Google account's Apps Script projects

    ```powershell
    clasp login
    ```

8. Enable the Apps Script API in [Google App Script User settings](https://script.google.com/home/usersettings)

---

## Optional commands

1. Install typescript compiler

    ```powershell
    npm install tsc
    ```

2. Install eslint in `.env`

    ```powershell
    npm install eslint
    ```

3. Create `.eselintconfig`

    ```powershell
    .\node_modules\.bin\eslint --init
    ```

4. Configure `.eslintconfig`

    ```powershell
    How would you like to use ESLint? · style       
    √ What type of modules does your project use? · commonjs
    √ Which framework does your project use? · none
    √ Does your project use TypeScript? · No
    √ Where does your code run? · browser (Node)
    √ How would you like to define a style for your project? · guide
    √ Which style guide do you want to follow? · standard
    √ What format do you want your config file to be in? · JSON
    ```

5. Install Prettier

    ```powershell
    npm install prettier
    ```

# Check Validity Of Digital Certificates
This simple Python script will get all your installed digital certificates and take a look in their validity, if it is missing 10 days to the due data, the script will tell you and generate 2 .xlsx files with all the caught information from the certificates.

Data that the script gonna get: Subject, Not Valid Before, Not Valid After.

To update that parameters, adjust the following code lines:

## Certificate Class
At first, adjust the Certificate class, adding or removing the wanted attributes:
![Class code](assets/img/class.png)
You can find that file at "entities/Certificado.py"

## Decode Script
Now, update the decoding script (located at main.py), adding or removing the wanted attributes, as described under:
![Class code](assets/img/config.png)

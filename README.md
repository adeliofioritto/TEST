oc create -f C:\P4C\Bot\TEST\secret.yaml
oc create -f C:\P4C\Bot\TEST\configmap.yaml
oc patch dc/report4c -p "{\"spec\": {\"strategy\": {\"\type": \"Recreate"}}}"
oc set volume deploymentconfig report4c --add --secret-name=report4c-secret --mount-path=/var/secret
oc set volume dc report4c --add --configmap-name=report4c-config -m /var/config![image](https://github.com/adeliofioritto/TEST/assets/6860894/658771ef-2dcc-425f-8d21-6b950c44e811)

docker pull wechaty/wechaty:latest

export WECHATY_LOG="verbose"
export WECHATY_PUPPET="wechaty-puppet-wechat"
export WECHATY_PUPPET_SERVER_PORT="9099"
export WECHATY_TOKEN="683e40b5-237f-4dc4-bed8-3918c0309f69"
export WECHATY_PUPPET_SERVICE_NO_TLS_INSECURE_SERVER="true"

# save login session
if [ ! -f "${WECHATY_TOKEN}.memory-card.json" ]; then
touch "${WECHATY_TOKEN}.memory-card.json"
fi

docker run -ti \
--name wechaty_puppet_service_token_gateway \
--rm \
-v "`pwd`/${WECHATY_TOKEN}.memory-card.json":"/wechaty/${WECHATY_TOKEN}.memory-card.json" \
-e WECHATY_LOG \
-e WECHATY_PUPPET \
-e WECHATY_PUPPET_SERVER_PORT \
-e WECHATY_PUPPET_SERVICE_NO_TLS_INSECURE_SERVER \
-e WECHATY_TOKEN \
-p "$WECHATY_PUPPET_SERVER_PORT:$WECHATY_PUPPET_SERVER_PORT" \
wechaty/wechaty:latest
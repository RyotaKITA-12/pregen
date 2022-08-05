docker compose build
docker compose run --rm client sh -c "yarn global add create-react-app && create-react-app ./ --template typescript"

# Use Node.js for build stage
FROM node:20 as build
WORKDIR /app

# Copy package.json and install dependencies
COPY package.json package-lock.json ./
RUN npm install

# Copy the rest of the code and build it
COPY . .
RUN npm run build

# Use Nginx for production
FROM nginx:latest
COPY --from=build /app/dist /usr/share/nginx/html
COPY ./nginx.conf /etc/nginx/nginx.conf

# Expose the correct port
EXPOSE 8080

# Start Nginx
CMD ["nginx", "-g", "daemon off;"]

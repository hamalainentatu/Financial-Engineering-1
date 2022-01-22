#TU-E2210 Financial Engineering 1, Aalto University
#Hedging Assignment for 5 cr course version
#Copyright Tatu Hämäläinen
#Updated 10.1.2022

### Delta Hedging ###
# 1. Hedging one ATM call option with maturity of 45 days
# 2. Hedging with different strike prices and rehedging frequencies (Table 1)
# 3. Testing the statistical significance of results of 2 by repeating above process to all 12 months/worksheets (Table 2)
# 4. Hedging with different initial days to maturity and rehedging frequencies (Table 3)
# 5. Hedging with different initial days to maturity and strike prices (Table 4)
# 6. Testing the statistical significance of results of 5 by repeating above process to all 12 months/worksheets (Table 5)

### Delta-vega Hedging ###
# 7. Delta-vega hedging an ATM call with maturity 45
# 8. Delta-vega hedging with different strike prices and rehedging frequencies (Table 6)
# 9. Testing the statistical significance of above results of 8 by repeating above process to 11 other months/worksheets (excl December) (Table 7)
# 10. Delta-vega hedging with respect to moneyness and initial days to maturity. All worksheets excl. December (Table 8)

# 11. Comparing the performance of strategies



library(readxl, quietly = TRUE)
library(plyr)
library(dplyr, quietly = TRUE)
library(tidyr, quietly = TRUE)
library(purrr, quietly = TRUE)
library(ggplot2, quietly = TRUE)
library(plotly, quietly = TRUE)
library(RQuantLib)
library(writexl)


######### DELTA HEDGING -- PARTS 1-6 ##############

### Reading the data ###

data <- read_excel("isx2010C.xls", sheet = 1) #January of 2010
data <- data %>% rename(T = names(data)[1]) %>% rename(S = names(data)[dim(data)[2]-2]) %>% rename(r = names(data)[dim(data)[2]-1]) #rename variables
data <- data[,1:dim(data)[2]-1] #drop the last column (date)

data1 <- data %>% pivot_longer(-c(T, r, S), names_to = "E", values_to = "Cobs" ) %>% filter(!is.na(Cobs)) %>% 
  mutate(E = map_dbl(E, as.numeric)) %>% mutate(r = r/100, T = T/252, S = S/1000, E = E/1000, Cobs = Cobs/1000) #converting to tidy data, and scaling the values

scaleOversized <- function(x){ #Scaling outliers
  if(x > 1) x <- x / 1000
  return(x)
}
data1 <- data1 %>% mutate(Cobs = map_dbl(Cobs, scaleOversized))




### Plotting stock price ###

# data1 %>%
#   ggplot(aes(x = desc(T)*252 , y = S)) +
#   geom_line() + geom_smooth(formula = y ~ poly(x,8) , se = TRUE) +
#   labs(title ="Stock prices", x = "Days to maturity", y = "Value")
# 


### Plotting call prices as a function of the time to maturity and the strike process ###

# plot <- data1 %>% plot_ly(x = ~(252*T), y = ~(1000*E), z = ~(1000*Cobs), type="mesh3d", intensity = c(0, 0.33, 0.66, 1), color = 0.67 , opacity=0.8 )%>% layout(title = "Observed Call prices", scene = list(
#   xaxis = list(title = "Time to maturity"),
#   yaxis = list(title = "Strike"),
#   zaxis = list(title = "Option price")
# ))
# plot


### Calculating implied volatility ###

vola <- function(Cobs, S, E, r, T){
  result <- try(EuropeanOptionImpliedVolatility(type="call", value=Cobs, underlying=S, strike=E, dividendYield=0, riskFreeRate=r, maturity=T, volatility=0.1), silent = TRUE)
  if(!is.numeric(result)){
    result <- NA
  }
  return(result)
}

#data2 <- data1 %>% mutate(volatility = pmap_dbl(data1, function(Cobs, S, E, r, T)vola(Cobs, S, E, r, T) ))
#print(data2)

data2 <- data1 %>% mutate(volatility = pmap_dbl(data1, vola))


#filter too large and small volatilities?
#data2 <- data2 %>% filter(volatility > 0 & volatility < 0.6)

# data2 %>% ggplot(aes(x=1000* E,y=volatility, color=252 * T, group = T))+
#   geom_line() + ggtitle("Volatility smile") +
#   xlab("Strike") + theme(legend.position = "none") +
#   guides(colour = guide_legend(override.aes = list(alpha = 1))) +
#   scale_color_gradient(low = "blue", high = "red")


### 1. Hedging one ATM call option with maturity of 45 days ###

Maturity <- 45
#Maturities <- c()

Spot45 <- data[dim(data)[1]-(Maturity-1), dim(data)[2]-1] #spot price at Maturity 45 days
Spot45

#choosing call with E=520, which is about ATM call (Spot price 516 at day 45)
Set1 <- filter(data2, E == 0.52) #
Set1

Mat45 <- Set1 %>% filter(T == 45/252)
S <- Mat45$S
E <- Mat45$E
r <- Mat45$r
T0 <- Mat45$T
sigma <- Mat45$volatility
Cobs <- Mat45$Cobs
N <- 1000

### Delta hedging ###

#initial portfolio

delta <- N * EuropeanOption(type="call", underlying=S, strike=E,
                            riskFreeRate=r, maturity=T0, volatility=sigma,
                            dividendYield=0)$delta

Nhedge <- round(delta)
Nhedge

N*Cobs-Nhedge*S
#dynamic hedging

C0 <- Cobs
S0 <- S
Asum <- 0
n <- 0
k <- 0 #rehedge when k is even, i.e. every other day

#write_xlsx(Set1, "test.xlsx")

for (time in 1:(Mat45$T*252 - 1)){
  i <- ((Mat45$T*252) - time)
  NewMat <- Set1 %>% filter(T == i/252)
  if (!empty(NewMat)) {
    if (!is.na(NewMat$volatility)) {
      C1 <- NewMat$Cobs
      dC <- N*(C1 - C0)
      C0 <- C1
      S1 <- NewMat$S
      dS <- Nhedge*(S1 - S0)
      S0 <- S1
      A <- (dC - dS)^2
      Asum <- Asum + A
      n <- n + 1
      k <- k + 1
      if ((k %% 2) == 0) { #rehedge when k is even, i.e. every other day
        Snew <- NewMat$S
        Enew <- NewMat$E
        rnew <- NewMat$r
        T0new <- NewMat$T
        sigmanew <- NewMat$volatility
        
        deltanew <- N * EuropeanOption(type="call", underlying=Snew, strike=Enew,
                                       riskFreeRate=rnew, maturity=T0new, volatility=sigmanew,
                                       dividendYield=0)$delta
        
        Nhedge <- round(deltanew)
        Nhedge
      }
    }
  }
  
}

TotalError <- Asum / n

N*C0-Nhedge*S0




### 2. Hedging with different strike prices and rehedging frequencies ###


strikes <- c(0.36, 0.40, 0.44, 0.48, 0.52, 0.56) #hedging calls with different strike prices
freq <- c(1,2,7) #hedging with different rehedging frequencies
results <- data.frame(matrix(vector(), 3, 6))

for (j in 1:length(strikes)) { #hedging calls with different strike prices
  Set1 <- filter(data2, E == strikes[j]) #filtering the data based on strike price
  
  Mat45 <- Set1 %>% filter(T == 45/252) #selecting initial day 45
  S <- Mat45$S
  E <- Mat45$E
  r <- Mat45$r
  T0 <- Mat45$T
  sigma <- Mat45$volatility
  Cobs <- Mat45$Cobs
  N <- 1000
  
  
  
  for (f in 1:length(freq)) { #hedging with different rehedging frequencies
    
    ### Delta hedging ###
    
    #initial portfolio
    
    delta <- N * EuropeanOption(type="call", underlying=S, strike=E,
                                riskFreeRate=r, maturity=T0, volatility=sigma,
                                dividendYield=0)$delta
    
    Nhedge <- round(delta)
    
    #dynamic hedging
    
    C0 <- Cobs
    S0 <- S
    Asum <- 0
    n <- 0
    k <- 0 #rehedging frequency
    
    for (time in 1:(Mat45$T*252 - 1)){
      i <- ((Mat45$T*252) - time)
      NewMat <- Set1 %>% filter(T == i/252)
      if (!empty(NewMat)) {
        if (!is.na(NewMat$volatility)) {
          C1 <- NewMat$Cobs
          dC <- N*(C1 - C0)
          C0 <- C1
          S1 <- NewMat$S
          dS <- Nhedge*(S1 - S0)
          S0 <- S1
          A <- (dC - dS)^2
          Asum <- Asum + A
          n <- n + 1
          k <- k + 1
          if ((k %% freq[f]) == 0) { #rehedge when remainder = 0
            Snew <- NewMat$S
            Enew <- NewMat$E
            rnew <- NewMat$r
            T0new <- NewMat$T
            sigmanew <- NewMat$volatility
            
            deltanew <- N * EuropeanOption(type="call", underlying=Snew, strike=Enew,
                                           riskFreeRate=rnew, maturity=T0new, volatility=sigmanew,
                                           dividendYield=0)$delta
            
            Nhedge <- round(deltanew)
            Nhedge
          }
        }
      }
      
    }
    
    TotalError <- Asum / n
    results[f,j] <- TotalError
    
  }
}


### 3. Testing the statistical significance of above results by repeating above process to all 12 months/worksheets ###

months <- c(1:12)

Moneyness <- c(0.8, 0.9, 1, 1.1) #rehedging with different moneyness values
freq <- c(1,2,7) #hedging with different rehedging frequencies

monthresults <- data.frame(matrix(vector(), length(months), length(Moneyness)*length(freq)))


for (l in 1:length(months)) {
  data <- read_excel("isx2010C.xls", sheet = l) 
  data <- data %>% rename(T = names(data)[1]) %>% rename(S = names(data)[dim(data)[2]-2]) %>% rename(r = names(data)[dim(data)[2]-1]) #rename variables
  data <- data[,1:dim(data)[2]-1] #drop the last column (date)
  data <- mutate_all(data, function(x) as.numeric(as.character(x)))
  
  data1 <- data %>% pivot_longer(-c(T, r, S), names_to = "E", values_to = "Cobs" ) %>% filter(!is.na(Cobs)) %>% 
    mutate(E = map_dbl(E, as.numeric)) %>% mutate(r = r/100, T = T/252, S = S/1000, E = E/1000, Cobs = Cobs/1000) #converting to tidy data, and scaling the values
  
  scaleOversized <- function(x){ #Scaling outliers
    if(x > 1) x <- x / 1000
    return(x)
  }
  data1 <- data1 %>% mutate(Cobs = map_dbl(Cobs, scaleOversized))
  
  
  ### Calculating implied volatility ###
  
  vola <- function(Cobs, S, E, r, T){
    result <- try(EuropeanOptionImpliedVolatility(type="call", value=Cobs, underlying=S, strike=E, dividendYield=0, riskFreeRate=r, maturity=T, volatility=0.1), silent = TRUE)
    if(!is.numeric(result)){
      result <- NA
    }
    return(result)
  }
  
  data2 <- data1 %>% mutate(volatility = pmap_dbl(data1, vola))
  
  
  ### Hedging one ATM call option with maturity of 45 days ###
  
  Maturity <- 45
  
  Spot45 <- data[dim(data)[1]-(Maturity-1), dim(data)[2]-1] / 1000 #spot price at Maturity 45 days
  r45 <- data[dim(data)[1]-(Maturity-1), dim(data)[2]] / 100 #r at Maturity 45 days
  
  
  calcstrike <- function(x, r, S, T) {
    res <- x*S*exp(r*T)
    return(res)
  }
  
  strikelist <- c()
  for (a in 1:length(Moneyness)) { #calculating corresponding strike prices from moneyness-vector
    res1 <- calcstrike(Moneyness[a], r45, Spot45, (Maturity/252)) * 100
    res2 <- 2*round(res1/2)/100 #round to nearest even integer
    strikelist[a] <- res2
  }
  
  
  results <- data.frame(matrix(vector(), length(freq), length(strikelist)))
  
  for (j in 1:length(strikelist)) { #hedging calls with different strike prices
    Set1 <- filter(data2, E == strikelist[j]) #filtering the data based on strike price
    if (!empty(Set1)) {
      Mat45 <- Set1 %>% filter(T == Maturity/252) #selecting initial day 45
      S <- Mat45$S
      E <- Mat45$E
      r <- Mat45$r
      T0 <- Mat45$T
      sigma <- Mat45$volatility
      Cobs <- Mat45$Cobs
      N <- 1000
      
      
      
      for (f in 1:length(freq)) { #hedging with different rehedging frequencies
        
        ### Delta hedging ###
        
        #initial portfolio
        
        if (is.na(sigma)) { #check if initial sigma is NA
          for (d in 1:(Maturity-1)) {
            Mat45 <- Set1 %>% filter(T == (Maturity-d)/252)
            S <- Mat45$S
            E <- Mat45$E
            r <- Mat45$r
            T0 <- Mat45$T
            sigma <- Mat45$volatility
            Cobs <- Mat45$Cobs
            
            if (!is.na(sigma)) {
              break
            }
          }
        }
        
        delta <- N * EuropeanOption(type="call", underlying=S, strike=E,
                                    riskFreeRate=r, maturity=T0, volatility=sigma,
                                    dividendYield=0)$delta
        
        Nhedge <- round(delta)
        
        #dynamic hedging
        
        C0 <- Cobs
        S0 <- S
        Asum <- 0
        n <- 0
        k <- 0 #rehedging frequency
        
        
        for (time in 1:(Mat45$T*252 - 1)){
          i <- ((Mat45$T*252) - time)
          NewMat <- Set1 %>% filter(T == i/252)
          if (!empty(NewMat)) {
            if (!is.na(NewMat$volatility)) {
              C1 <- NewMat$Cobs
              dC <- N*(C1 - C0)
              C0 <- C1
              S1 <- NewMat$S
              dS <- Nhedge*(S1 - S0)
              S0 <- S1
              A <- (dC - dS)^2
              Asum <- Asum + A
              n <- n + 1
              k <- k + 1
              if ((k %% freq[f]) == 0) { #rehedge when remainder = 0
                Snew <- NewMat$S
                Enew <- NewMat$E
                rnew <- NewMat$r
                T0new <- NewMat$T
                sigmanew <- NewMat$volatility
                
                deltanew <- N * EuropeanOption(type="call", underlying=Snew, strike=Enew,
                                               riskFreeRate=rnew, maturity=T0new, volatility=sigmanew,
                                               dividendYield=0)$delta
                
                Nhedge <- round(deltanew)
              }
            }
          }
          
        }
        
        TotalError <- Asum / n
        results[f,j] <- TotalError
        
      }
    }
    
  }
  
  for (i in 1:(length(strikelist))) {
    for (j in 1:(length(freq))) {
      if (!is.na(results[j,i])) {
        monthresults[l,(i+(j-1)*length(strikelist))] <- results[j,i]
      } 
    }
    
  }
}

averages1 <- colMeans(monthresults, na.rm = TRUE) # average of each month
sd1 <- sapply(monthresults, sd, na.rm = TRUE) # sd of each month






### 4. Hedging with different initial days to maturity and rehedging frequencies ###

maturities <- c(85, 60, 45, 30, 15) #hedging calls with different initial day
freq <- c(1,2,7) #hedging with different rehedging frequencies
results <- data.frame(matrix(vector(), 3, 5))

for (j in 1:length(maturities)) { #hedging calls with different initial day
  Set1 <- filter(data2, E == 0.48) #filtering the data based on strike price = 0.52
  
  Mat45 <- Set1 %>% filter(T == maturities[j]/252) #selecting initial day 45
  S <- Mat45$S
  E <- Mat45$E
  r <- Mat45$r
  T0 <- Mat45$T
  sigma <- Mat45$volatility
  Cobs <- Mat45$Cobs
  N <- 1000
  
  
  
  for (f in 1:length(freq)) { #hedging with different rehedging frequencies
    
    ### Delta hedging ###
    
    #initial portfolio
    
    delta <- N * EuropeanOption(type="call", underlying=S, strike=E,
                                riskFreeRate=r, maturity=T0, volatility=sigma,
                                dividendYield=0)$delta
    
    Nhedge <- round(delta)
    
    #dynamic hedging
    
    C0 <- Cobs
    S0 <- S
    Asum <- 0
    n <- 0
    k <- 0 #rehedging frequency
    
    for (time in 1:(Mat45$T*252 - 1)){
      i <- ((Mat45$T*252) - time)
      NewMat <- Set1 %>% filter(T == i/252)
      if (!empty(NewMat)) {
        if (!is.na(NewMat$volatility)) {
          C1 <- NewMat$Cobs
          dC <- N*(C1 - C0)
          C0 <- C1
          S1 <- NewMat$S
          dS <- Nhedge*(S1 - S0)
          S0 <- S1
          A <- (dC - dS)^2
          Asum <- Asum + A
          n <- n + 1
          k <- k + 1
          if ((k %% freq[f]) == 0) { #rehedge when remainder = 0
            Snew <- NewMat$S
            Enew <- NewMat$E
            rnew <- NewMat$r
            T0new <- NewMat$T
            sigmanew <- NewMat$volatility
            
            deltanew <- N * EuropeanOption(type="call", underlying=Snew, strike=Enew,
                                           riskFreeRate=rnew, maturity=T0new, volatility=sigmanew,
                                           dividendYield=0)$delta
            
            Nhedge <- round(deltanew)
            Nhedge
          }
        }
      }
      
    }
    
    TotalError <- Asum / n
    results[f,j] <- TotalError
    
  }
}


### 5. Hedging with different initial days to maturity and strike prices ###

strikes <- c(0.36, 0.40, 0.44, 0.48, 0.52, 0.56) #hedging calls with different strike prices
maturities <- c(85, 60, 45, 30, 15) #hedging calls with different initial day
results <- data.frame(matrix(vector(), length(maturities), length(strikes)))

for (j in 1:length(strikes)) { #hedging calls with different strike prices
  Set1 <- filter(data2, E == strikes[j]) #filtering the data based on strike price
  
  for (f in 1:length(maturities)) { #hedging with different rehedging frequencies
    
    ### Delta hedging ###
    
    #initial portfolio
    
    Mat45 <- Set1 %>% filter(T == maturities[f]/252) #selecting initial day 45
    S <- Mat45$S
    E <- Mat45$E
    r <- Mat45$r
    T0 <- Mat45$T
    sigma <- Mat45$volatility
    Cobs <- Mat45$Cobs
    N <- 1000
    
    for (d in 1:(maturities[f]-1)) { #check if initial row is empty or sigma is NA
      if (empty(Mat45)) {
        Mat45 <- Set1 %>% filter(T == (maturities[f]-d)/252)
        S <- Mat45$S
        E <- Mat45$E
        r <- Mat45$r
        T0 <- Mat45$T
        sigma <- Mat45$volatility
        Cobs <- Mat45$Cobs
      } else if (is.na(sigma)) {
        Mat45 <- Set1 %>% filter(T == (maturities[f]-d)/252)
        S <- Mat45$S
        E <- Mat45$E
        r <- Mat45$r
        T0 <- Mat45$T
        sigma <- Mat45$volatility
        Cobs <- Mat45$Cobs
      } else {
        break
      }
    }
      
    
    delta <- N * EuropeanOption(type="call", underlying=S, strike=E,
                                riskFreeRate=r, maturity=T0, volatility=sigma,
                                dividendYield=0)$delta
    
    Nhedge <- round(delta)
    
    #dynamic hedging
    
    C0 <- Cobs
    S0 <- S
    Asum <- 0
    n <- 0
    k <- 0 #rehedging frequency
    
    for (time in 1:(Mat45$T*252 - 1)){
      i <- ((Mat45$T*252) - time)
      NewMat <- Set1 %>% filter(T == i/252)
      if (!empty(NewMat)) {
        if (!is.na(NewMat$volatility)) {
          C1 <- NewMat$Cobs
          dC <- N*(C1 - C0)
          C0 <- C1
          S1 <- NewMat$S
          dS <- Nhedge*(S1 - S0)
          S0 <- S1
          A <- (dC - dS)^2
          Asum <- Asum + A
          n <- n + 1
          k <- k + 1
          if (((k %% 2) == 0)) { #rehedge when remainder = 0, i.e. every other day
            Snew <- NewMat$S
            Enew <- NewMat$E
            rnew <- NewMat$r
            T0new <- NewMat$T
            sigmanew <- NewMat$volatility
            
            deltanew <- N * EuropeanOption(type="call", underlying=Snew, strike=Enew,
                                           riskFreeRate=rnew, maturity=T0new, volatility=sigmanew,
                                           dividendYield=0)$delta
            
            Nhedge <- round(deltanew)
            Nhedge
          }
        }
      }
      
    }
    
    TotalError <- Asum / n
    results[f,j] <- TotalError
    
  }
}



### 6. Testing the statistical significance of results of 5 by repeating above process to all 12 months/worksheets ####

months <- c(1:12)

Moneyness <- c(0.8, 0.9, 1, 1.1) #hedging with different moneyness values
maturities <- c(85, 60, 45, 30, 15) #hedging calls with different initial day

monthresults <- data.frame(matrix(vector(), length(months), length(Moneyness)*length(maturities)))

for (l in 1:length(months)) {
  
  data <- read_excel("isx2010C.xls", sheet = l) 
  data <- data %>% rename(T = names(data)[1]) %>% rename(S = names(data)[dim(data)[2]-2]) %>% rename(r = names(data)[dim(data)[2]-1]) #rename variables
  data <- data[,1:dim(data)[2]-1] #drop the last column (date)
  data <- mutate_all(data, function(x) as.numeric(as.character(x)))
  
  data1 <- data %>% pivot_longer(-c(T, r, S), names_to = "E", values_to = "Cobs" ) %>% filter(!is.na(Cobs)) %>% 
    mutate(E = map_dbl(E, as.numeric)) %>% mutate(r = r/100, T = T/252, S = S/1000, E = E/1000, Cobs = Cobs/1000) #converting to tidy data, and scaling the values
  
  scaleOversized <- function(x){ #Scaling outliers
    if(x > 1) x <- x / 1000
    return(x)
  }
  data1 <- data1 %>% mutate(Cobs = map_dbl(Cobs, scaleOversized))
  
  
  ### Calculating implied volatility ###
  
  vola <- function(Cobs, S, E, r, T){
    result <- try(EuropeanOptionImpliedVolatility(type="call", value=Cobs, underlying=S, strike=E, dividendYield=0, riskFreeRate=r, maturity=T, volatility=0.1), silent = TRUE)
    if(!is.numeric(result)){
      result <- NA
    }
    return(result)
  }
  
  data2 <- data1 %>% mutate(volatility = pmap_dbl(data1, vola))
  
  ### Dynamic Hedging ###
  
  results <- data.frame(matrix(vector(), length(maturities), length(Moneyness)))
  
  for (j in 1:length(Moneyness)) { #hedging calls with different strike prices
      
      for (f in 1:length(maturities)) { #hedging with different rehedging frequencies
        
        ### Strike prices ###
        
        Maturity <- maturities[f]
        
        Spot45 <- data[dim(data)[1]-(Maturity-1), dim(data)[2]-1] / 1000 #spot price at Maturity 45 days
        r45 <- data[dim(data)[1]-(Maturity-1), dim(data)[2]] / 100 #r at Maturity 45 days
        
        calcstrike <- function(x, r, S, T) { #calculate strike price from moneyness
          res <- x*S*exp(r*T)
          return(res)
        }
        
        strikelist <- c()
        for (a in 1:length(Moneyness)) { #calculating corresponding strike prices from moneyness-vector
          res1 <- calcstrike(Moneyness[a], r45, Spot45, (Maturity/252)) * 100
          res2 <- 2*round(res1/2)/100 #round to nearest even integer
          strikelist[a] <- res2
        }
        
        #initial portfolio
        
        Set1 <- filter(data2, E == strikelist[j]) #filtering the data based on strike price
        
        if (!empty(Set1)) {
          
          Mat45 <- Set1 %>% filter(T == maturities[f]/252) #selecting initial day 45
          S <- Mat45$S
          E <- Mat45$E
          r <- Mat45$r
          T0 <- Mat45$T
          sigma <- Mat45$volatility
          Cobs <- Mat45$Cobs
          N <- 1000
          
          for (d in 1:(maturities[f]-1)) { #check if initial row is empty or sigma is NA
            if (empty(Mat45)) {
              Mat45 <- Set1 %>% filter(T == (maturities[f]-d)/252)
              S <- Mat45$S
              E <- Mat45$E
              r <- Mat45$r
              T0 <- Mat45$T
              sigma <- Mat45$volatility
              Cobs <- Mat45$Cobs
            } else if (is.na(sigma)) {
              Mat45 <- Set1 %>% filter(T == (maturities[f]-d)/252)
              S <- Mat45$S
              E <- Mat45$E
              r <- Mat45$r
              T0 <- Mat45$T
              sigma <- Mat45$volatility
              Cobs <- Mat45$Cobs
            } else {
              break
            }
          }
          
          
          delta <- N * EuropeanOption(type="call", underlying=S, strike=E,
                                      riskFreeRate=r, maturity=T0, volatility=sigma,
                                      dividendYield=0)$delta
          
          Nhedge <- round(delta)
          
          #dynamic hedging
          
          C0 <- Cobs
          S0 <- S
          Asum <- 0
          n <- 0
          k <- 0 #rehedging frequency
          
          for (time in 1:(Mat45$T*252 - 1)){
            i <- ((Mat45$T*252) - time)
            NewMat <- Set1 %>% filter(T == i/252)
            if (!empty(NewMat)) {
              if (!is.na(NewMat$volatility)) {
                C1 <- NewMat$Cobs
                dC <- N*(C1 - C0)
                C0 <- C1
                S1 <- NewMat$S
                dS <- Nhedge*(S1 - S0)
                S0 <- S1
                A <- (dC - dS)^2
                Asum <- Asum + A
                n <- n + 1
                k <- k + 1
                if (((k %% 2) == 0)) { #rehedge when remainder = 0, i.e. every other day
                  Snew <- NewMat$S
                  Enew <- NewMat$E
                  rnew <- NewMat$r
                  T0new <- NewMat$T
                  sigmanew <- NewMat$volatility
                  
                  deltanew <- N * EuropeanOption(type="call", underlying=Snew, strike=Enew,
                                                 riskFreeRate=rnew, maturity=T0new, volatility=sigmanew,
                                                 dividendYield=0)$delta
                  
                  Nhedge <- round(deltanew)
                  Nhedge
                }
              }
            }
            
          }
          TotalError <- Asum / n
          results[f,j] <- TotalError
          
        }
      }
  }
  
  for (i in 1:(length(Moneyness))) {
    for (j in 1:(length(maturities))) {
      if (!is.na(results[j,i])) {
        monthresults[l,(i+(j-1)*length(Moneyness))] <- results[j,i]
      } 
    }
    
  }
}

averages2 <- colMeans(monthresults, na.rm = TRUE) # average of each month
sd2 <- sapply(monthresults, sd, na.rm = TRUE) # sd of each month




########################## DELTA-VEGA HEDGING #########################

###### 7. Delta-vega hedging an ATM call with maturity 45 ####### 

#### reading the January data ###

data <- read_excel("isx2010C.xls", sheet = 1) 
data <- data %>% rename(T = names(data)[1]) %>% rename(S = names(data)[dim(data)[2]-2]) %>% rename(r = names(data)[dim(data)[2]-1]) #rename variables
dates1 <- data[,dim(data)[2]]
data <- data[,1:dim(data)[2]-1] #drop the last column (date)
data <- mutate_all(data, function(x) as.numeric(as.character(x)))
data["date"] <- dates1
data <- data %>% rename(date = names(data)[dim(data)[2]])
data <- data %>% mutate(date=as.Date(date, format = "%d.%m.%Y"))

data1 <- data %>% pivot_longer(-c(T, r, S, date), names_to = "E", values_to = "Cobs" ) %>% filter(!is.na(Cobs)) %>% 
  mutate(E = map_dbl(E, as.numeric)) %>% mutate(r = r/100, T = T/252, S = S/1000, E = E/1000, Cobs = Cobs/1000) #converting to tidy data, and scaling the values

scaleOversized <- function(x){ #Scaling outliers
  if(x > 1) x <- x / 1000
  return(x)
}
data1 <- data1 %>% mutate(Cobs = map_dbl(Cobs, scaleOversized))


### Calculating implied volatility ###

vola <- function(Cobs, S, E, r, T){
  result <- try(EuropeanOptionImpliedVolatility(type="call", value=Cobs, underlying=S, strike=E, dividendYield=0, riskFreeRate=r, maturity=T, volatility=0.1), silent = TRUE)
  if(!is.numeric(result)){
    result <- NA
  }
  return(result)
}

data2 <- data1 %>% mutate(volatility = pmap_dbl(data1[, c('Cobs', 'S', 'E', 'r', 'T')], vola))

#### reading the February data ###

datanew <- read_excel("isx2010C.xls", sheet = 2) 
datanew <- datanew %>% rename(T = names(datanew)[1]) %>% rename(S = names(datanew)[dim(datanew)[2]-2]) %>% rename(r = names(datanew)[dim(datanew)[2]-1]) #rename variables
dates2 <- datanew[,dim(datanew)[2]]
datanew <- datanew[,1:dim(datanew)[2]-1] #drop the last column (date)
datanew <- mutate_all(datanew, function(x) as.numeric(as.character(x)))
datanew["date"] <- dates2
datanew <- datanew %>% mutate(date=as.Date(date, format = "%d.%m.%Y"))

datanew1 <- datanew %>% pivot_longer(-c(T, r, S, date), names_to = "E", values_to = "Cobs" ) %>% filter(!is.na(Cobs)) %>% 
  mutate(E = map_dbl(E, as.numeric)) %>% mutate(r = r/100, T = T/252, S = S/1000, E = E/1000, Cobs = Cobs/1000) #converting to tidy data, and scaling the values

scaleOversized <- function(x){ #Scaling outliers
  if(x > 1) x <- x / 1000
  return(x)
}
datanew1 <- datanew1 %>% mutate(Cobs = map_dbl(Cobs, scaleOversized))


### Calculating implied volatility ###

vola <- function(Cobs, S, E, r, T){
  result <- try(EuropeanOptionImpliedVolatility(type="call", value=Cobs, underlying=S, strike=E, dividendYield=0, riskFreeRate=r, maturity=T, volatility=0.1), silent = TRUE)
  if(!is.numeric(result)){
    result <- NA
  }
  return(result)
}

datanew2 <- datanew1 %>% mutate(volatility = pmap_dbl(datanew1[, c('Cobs', 'S', 'E', 'r', 'T')], vola))

# calculating the days to maturity of a replicating option

Maturity1 <- 45
Spot45 <- data[dim(data)[1]-(Maturity1-1), dim(data)[2]-2] / 1000 #spot price at Maturity 45 days
date45 <- data[dim(data)[1]-(Maturity1-1), dim(data)[2]]

Maturity2 <- data.frame(datanew) %>% filter(date == date45)
Maturity2 <- Maturity2$T

Set1 <- filter(data2, E == 0.52) #filtering the data based on strike price
Set2 <- filter(datanew2, E == 0.52)

if (!empty(Set1) && !empty(Set2)) {
  
  Mat45 <- Set1 %>% filter(T == Maturity1/252) #selecting initial day 45
  S <- Mat45$S
  E <- Mat45$E
  r <- Mat45$r
  T0 <- Mat45$T
  sigma <- Mat45$volatility
  Cobs <- Mat45$Cobs
  N <- 1000
  
  MatRep <- Set2 %>% filter(T == Maturity2/252) #selecting initial day that has same date than Maturity1
  SRep <- MatRep$S
  ERep <- MatRep$E
  rRep <- MatRep$r
  T0Rep <- MatRep$T
  sigmaRep <- MatRep$volatility
  CobsRep <- MatRep$Cobs
  
  for (d in 1:(Maturity1-1)) { #check if initial rows are empty or sigma is NA
    if (empty(Mat45) || empty(MatRep)) {
      Mat45 <- Set1 %>% filter(T == (Maturity1-d)/252)
      S <- Mat45$S
      E <- Mat45$E
      r <- Mat45$r
      T0 <- Mat45$T
      sigma <- Mat45$volatility
      Cobs <- Mat45$Cobs
      
      MatRep <- Set2 %>% filter(T == Maturity2-d/252)
      SRep <- MatRep$S
      ERep <- MatRep$E
      rRep <- MatRep$r
      T0Rep <- MatRep$T
      sigmaRep <- MatRep$volatility
      CobsRep <- MatRep$Cobs
      
    } else if (is.na(sigma) || is.na(sigmaRep)) {
      Mat45 <- Set1 %>% filter(T == (Maturity1-d)/252)
      S <- Mat45$S
      E <- Mat45$E
      r <- Mat45$r
      T0 <- Mat45$T
      sigma <- Mat45$volatility
      Cobs <- Mat45$Cobs
      
      MatRep <- Set2 %>% filter(T == Maturity2-d/252)
      SRep <- MatRep$S
      ERep <- MatRep$E
      rRep <- MatRep$r
      T0Rep <- MatRep$T
      sigmaRep <- MatRep$volatility
      CobsRep <- MatRep$Cobs
      
    } else {
      break
    }
  }
  
  #initial portfolio
  
  delta45 <- EuropeanOption(type="call", underlying=S, strike=E,
                              riskFreeRate=r, maturity=T0, volatility=sigma,
                              dividendYield=0)$delta
  
  deltaRep <- EuropeanOption(type="call", underlying=SRep, strike=ERep,
                                riskFreeRate=rRep, maturity=T0Rep, volatility=sigmaRep,
                                dividendYield=0)$delta
  
  vega45 <- EuropeanOption(type="call", underlying=S, strike=E,
                            riskFreeRate=r, maturity=T0, volatility=sigma,
                            dividendYield=0)$vega
  
  vegaRep <- EuropeanOption(type="call", underlying=SRep, strike=ERep,
                             riskFreeRate=rRep, maturity=T0Rep, volatility=sigmaRep,
                             dividendYield=0)$vega
  
  NHedgeStock <- round(N * (delta45 - (vega45/vegaRep)*deltaRep)) # number of indice to hold
  
  NHedgeRep <- round(N * (vega45/vegaRep)) # amount of replicating options to hold
  
  N * Cobs - NHedgeStock * S - NHedgeRep * CobsRep # value of the initial portfolio
  
  #dynamic hedging
  
  C0 <- Cobs
  S0 <- S
  C0Rep <- CobsRep
  S0Rep <- SRep
  
  Asum <- 0
  n <- 0
  k <- 0 #rehedging frequency
  
  for (time in 1:(Mat45$T*252 - 1)){
    i <- ((Mat45$T*252) - time)
    iRep <- ((MatRep$T*252) - time)
    NewMat <- Set1 %>% filter(T == i/252)
    NewMatRep <- Set2 %>% filter(T == iRep/252)
    if (!empty(NewMat) && !empty(NewMatRep)) {
      if (!is.na(NewMat$volatility) && !is.na(NewMatRep$volatility)) {
        C1 <- NewMat$Cobs
        dC <- N*(C1 - C0)
        C0 <- C1
        S1 <- NewMat$S
        dS <- NHedgeStock*(S1 - S0)
        S0 <- S1
        C1Rep <- NewMatRep$Cobs
        dCRep <- NHedgeRep*(C1Rep - C0Rep)
        C0Rep <- C1Rep
        A <- (dC - dS - dCRep)^2 
        Asum <- Asum + A
        n <- n + 1
        k <- k + 1
        if (((k %% 2) == 0)) { #rehedge when remainder = 0, i.e. every other day
          Snew <- NewMat$S
          Enew <- NewMat$E
          rnew <- NewMat$r
          T0new <- NewMat$T
          sigmanew <- NewMat$volatility
          
          SnewRep <- NewMatRep$S
          EnewRep <- NewMatRep$E
          rnewRep <- NewMatRep$r
          T0newRep <- NewMatRep$T
          sigmanewRep <- NewMatRep$volatility

          deltanew <- EuropeanOption(type="call", underlying=Snew, strike=Enew,
                                         riskFreeRate=rnew, maturity=T0new, volatility=sigmanew,
                                         dividendYield=0)$delta
          
          deltaRep <- EuropeanOption(type="call", underlying=SnewRep, strike=EnewRep,
                                     riskFreeRate=rnewRep, maturity=T0newRep, volatility=sigmanewRep,
                                     dividendYield=0)$delta
          
          vega45 <- EuropeanOption(type="call", underlying=S, strike=E,
                                   riskFreeRate=r, maturity=T0, volatility=sigma,
                                   dividendYield=0)$vega
          
          vegaRep <- EuropeanOption(type="call", underlying=SnewRep, strike=EnewRep,
                                    riskFreeRate=rnewRep, maturity=T0newRep, volatility=sigmanewRep,
                                    dividendYield=0)$vega
          
          NHedgeStock <- round(N * (delta45 - (vega45/vegaRep)*deltaRep)) # number of stocks to hold
          
          NHedgeRep <- round(N * (vega45/vegaRep)) # amount of replicating options to hold
          
        }
      }
    }
  }
  TotalError <- Asum / n
  
}




### 8. Delta-vega hedging with different strike prices and rehedging frequencies ###
  
strikes <- c(0.36, 0.40, 0.44, 0.48, 0.52, 0.56) #hedging calls with different strike prices
freq <- c(1,2,7) #hedging with different rehedging frequencies
results <- data.frame(matrix(vector(), 3, 6))

for (j in 1:length(strikes)) { #hedging calls with different strike prices
  
  Maturity1 <- 45
  Spot45 <- data[dim(data)[1]-(Maturity1-1), dim(data)[2]-2] / 1000 #spot price at Maturity 45 days
  date45 <- data[dim(data)[1]-(Maturity1-1), dim(data)[2]]
  
  Maturity2 <- data.frame(datanew) %>% filter(date == date45)
  Maturity2 <- Maturity2$T
  
  Set1 <- filter(data2, E == strikes[j]) #filtering the data based on strike price
  Set2 <- filter(datanew2, E == strikes[j])
  
  if (!empty(Set1) && !empty(Set2)) {
    
    Mat45 <- Set1 %>% filter(T == Maturity1/252) #selecting initial day 45
    S <- Mat45$S
    E <- Mat45$E
    r <- Mat45$r
    T0 <- Mat45$T
    sigma <- Mat45$volatility
    Cobs <- Mat45$Cobs
    N <- 1000
    
    MatRep <- Set2 %>% filter(T == Maturity2/252) #selecting initial day that has same date than Maturity1
    SRep <- MatRep$S
    ERep <- MatRep$E
    rRep <- MatRep$r
    T0Rep <- MatRep$T
    sigmaRep <- MatRep$volatility
    CobsRep <- MatRep$Cobs
    
    for (d in 1:(Maturity1-1)) { #check if initial rows are empty or sigma is NA
      if (empty(Mat45) || empty(MatRep)) {
        Mat45 <- Set1 %>% filter(T == (Maturity1-d)/252)
        S <- Mat45$S
        E <- Mat45$E
        r <- Mat45$r
        T0 <- Mat45$T
        sigma <- Mat45$volatility
        Cobs <- Mat45$Cobs
        
        MatRep <- Set2 %>% filter(T == Maturity2-d/252)
        SRep <- MatRep$S
        ERep <- MatRep$E
        rRep <- MatRep$r
        T0Rep <- MatRep$T
        sigmaRep <- MatRep$volatility
        CobsRep <- MatRep$Cobs
        
      } else if (is.na(sigma) || is.na(sigmaRep)) {
        Mat45 <- Set1 %>% filter(T == (Maturity1-d)/252)
        S <- Mat45$S
        E <- Mat45$E
        r <- Mat45$r
        T0 <- Mat45$T
        sigma <- Mat45$volatility
        Cobs <- Mat45$Cobs
        
        MatRep <- Set2 %>% filter(T == Maturity2-d/252)
        SRep <- MatRep$S
        ERep <- MatRep$E
        rRep <- MatRep$r
        T0Rep <- MatRep$T
        sigmaRep <- MatRep$volatility
        CobsRep <- MatRep$Cobs
        
      } else {
        break
      }
    }
    
    
    
    for (f in 1:length(freq)) { #hedging with different rehedging frequencies
      
      #initial portfolio
      
      delta45 <- EuropeanOption(type="call", underlying=S, strike=E,
                                riskFreeRate=r, maturity=T0, volatility=sigma,
                                dividendYield=0)$delta
      
      deltaRep <- EuropeanOption(type="call", underlying=SRep, strike=ERep,
                                 riskFreeRate=rRep, maturity=T0Rep, volatility=sigmaRep,
                                 dividendYield=0)$delta
      
      vega45 <- EuropeanOption(type="call", underlying=S, strike=E,
                               riskFreeRate=r, maturity=T0, volatility=sigma,
                               dividendYield=0)$vega
      
      vegaRep <- EuropeanOption(type="call", underlying=SRep, strike=ERep,
                                riskFreeRate=rRep, maturity=T0Rep, volatility=sigmaRep,
                                dividendYield=0)$vega
      
      NHedgeStock <- round(N * (delta45 - (vega45/vegaRep)*deltaRep)) # number of indice to hold
      
      NHedgeRep <- round(N * (vega45/vegaRep)) # amount of replicating options to hold
      
      N * Cobs - NHedgeStock * S - NHedgeRep * CobsRep # value of the initial portfolio
      
      # Dynamic hedging
      C0 <- Cobs
      S0 <- S
      C0Rep <- CobsRep
      S0Rep <- SRep
      
      Asum <- 0
      n <- 0
      k <- 0 #rehedging frequency
      
      for (time in 1:(Mat45$T*252 - 1)){
        i <- ((Mat45$T*252) - time)
        iRep <- ((MatRep$T*252) - time)
        NewMat <- Set1 %>% filter(T == i/252)
        NewMatRep <- Set2 %>% filter(T == iRep/252)
        if (!empty(NewMat) && !empty(NewMatRep)) {
          if (!is.na(NewMat$volatility) && !is.na(NewMatRep$volatility)) {
            C1 <- NewMat$Cobs
            dC <- N*(C1 - C0)
            C0 <- C1
            S1 <- NewMat$S
            dS <- NHedgeStock*(S1 - S0)
            S0 <- S1
            C1Rep <- NewMatRep$Cobs
            dCRep <- NHedgeRep*(C1Rep - C0Rep)
            C0Rep <- C1Rep
            A <- (dC - dS - dCRep)^2 
            Asum <- Asum + A
            n <- n + 1
            k <- k + 1
            if (((k %% freq[f]) == 0)) { #rehedge when remainder = 0, i.e. every other day
              Snew <- NewMat$S
              Enew <- NewMat$E
              rnew <- NewMat$r
              T0new <- NewMat$T
              sigmanew <- NewMat$volatility
              
              SnewRep <- NewMatRep$S
              EnewRep <- NewMatRep$E
              rnewRep <- NewMatRep$r
              T0newRep <- NewMatRep$T
              sigmanewRep <- NewMatRep$volatility
              
              deltanew <- EuropeanOption(type="call", underlying=Snew, strike=Enew,
                                         riskFreeRate=rnew, maturity=T0new, volatility=sigmanew,
                                         dividendYield=0)$delta
              
              deltaRep <- EuropeanOption(type="call", underlying=SnewRep, strike=EnewRep,
                                         riskFreeRate=rnewRep, maturity=T0newRep, volatility=sigmanewRep,
                                         dividendYield=0)$delta
              
              vega45 <- EuropeanOption(type="call", underlying=S, strike=E,
                                       riskFreeRate=r, maturity=T0, volatility=sigma,
                                       dividendYield=0)$vega
              
              vegaRep <- EuropeanOption(type="call", underlying=SnewRep, strike=EnewRep,
                                        riskFreeRate=rnewRep, maturity=T0newRep, volatility=sigmanewRep,
                                        dividendYield=0)$vega
              
              NHedgeStock <- round(N * (delta45 - (vega45/vegaRep)*deltaRep)) # number of stocks to hold
              
              NHedgeRep <- round(N * (vega45/vegaRep)) # amount of replicating options to hold
              
            }
          }
        }
      }
      TotalError <- Asum / n
      results[f,j] <- TotalError
    }
  }
}

### 9. Testing the statistical significance of above results of 8 by repeating above process to 11 other months/worksheets (excl December) ###


months <- c(1:11)

Moneyness <- c(0.8, 0.9, 1, 1.1) #rehedging with different moneyness values
freq <- c(1,2,7) #hedging with different rehedging frequencies

monthresults <- data.frame(matrix(vector(), length(months), length(Moneyness)*length(freq)))

for (l in 1:length(months)) {
  
  data <- read_excel("isx2010C.xls", sheet = l) 
  data <- data %>% rename(T = names(data)[1]) %>% rename(S = names(data)[dim(data)[2]-2]) %>% rename(r = names(data)[dim(data)[2]-1]) #rename variables
  dates1 <- data[,dim(data)[2]]
  data <- data[,1:dim(data)[2]-1] #drop the last column (date)
  data <- mutate_all(data, function(x) as.numeric(as.character(x)))
  data["date"] <- dates1
  data <- data %>% rename(date = names(data)[dim(data)[2]])
  data <- data %>% mutate(date=as.Date(date, format = "%d.%m.%Y"))
  
  data1 <- data %>% pivot_longer(-c(T, r, S, date), names_to = "E", values_to = "Cobs" ) %>% filter(!is.na(Cobs)) %>% 
    mutate(E = map_dbl(E, as.numeric)) %>% mutate(r = r/100, T = T/252, S = S/1000, E = E/1000, Cobs = Cobs/1000) #converting to tidy data, and scaling the values
  
  scaleOversized <- function(x){ #Scaling outliers
    if(x > 1) x <- x / 1000
    return(x)
  }
  data1 <- data1 %>% mutate(Cobs = map_dbl(Cobs, scaleOversized))
  
  
  ### Calculating implied volatility ###
  
  vola <- function(Cobs, S, E, r, T){
    result <- try(EuropeanOptionImpliedVolatility(type="call", value=Cobs, underlying=S, strike=E, dividendYield=0, riskFreeRate=r, maturity=T, volatility=0.1), silent = TRUE)
    if(!is.numeric(result)){
      result <- NA
    }
    return(result)
  }
  
  data2 <- data1 %>% mutate(volatility = pmap_dbl(data1[, c('Cobs', 'S', 'E', 'r', 'T')], vola))
  
  #### reading the February data ###
  
  datanew <- read_excel("isx2010C.xls", sheet = (l+1)) 
  datanew <- datanew %>% rename(T = names(datanew)[1]) %>% rename(S = names(datanew)[dim(datanew)[2]-2]) %>% rename(r = names(datanew)[dim(datanew)[2]-1]) #rename variables
  dates2 <- datanew[,dim(datanew)[2]]
  datanew <- datanew[,1:dim(datanew)[2]-1] #drop the last column (date)
  datanew <- mutate_all(datanew, function(x) as.numeric(as.character(x)))
  datanew["date"] <- dates2
  datanew <- datanew %>% mutate(date=as.Date(date, format = "%d.%m.%Y"))
  
  datanew1 <- datanew %>% pivot_longer(-c(T, r, S, date), names_to = "E", values_to = "Cobs" ) %>% filter(!is.na(Cobs)) %>% 
    mutate(E = map_dbl(E, as.numeric)) %>% mutate(r = r/100, T = T/252, S = S/1000, E = E/1000, Cobs = Cobs/1000) #converting to tidy data, and scaling the values
  
  scaleOversized <- function(x){ #Scaling outliers
    if(x > 1) x <- x / 1000
    return(x)
  }
  datanew1 <- datanew1 %>% mutate(Cobs = map_dbl(Cobs, scaleOversized))
  
  
  ### Calculating implied volatility ###
  
  vola <- function(Cobs, S, E, r, T){
    result <- try(EuropeanOptionImpliedVolatility(type="call", value=Cobs, underlying=S, strike=E, dividendYield=0, riskFreeRate=r, maturity=T, volatility=0.1), silent = TRUE)
    if(!is.numeric(result)){
      result <- NA
    }
    return(result)
  }
  
  datanew2 <- datanew1 %>% mutate(volatility = pmap_dbl(datanew1[, c('Cobs', 'S', 'E', 'r', 'T')], vola))
  
  # calculating the strike prices from moneyness
  
  Maturity1 <- 45
  
  Spot45 <- data[dim(data)[1]-(Maturity1-1), dim(data)[2]-2] / 1000 #spot price at Maturity 45 days
  r45 <- data[dim(data)[1]-(Maturity1-1), dim(data)[2]-1] / 100 #r at Maturity 45 days
  
  calcstrike <- function(x, r, S, T) {
    res <- x*S*exp(r*T)
    return(res)
  }
  
  strikelist <- c()
  for (a in 1:length(Moneyness)) { #calculating corresponding strike prices from moneyness-vector
    res1 <- calcstrike(Moneyness[a], r45, Spot45, (Maturity1/252)) * 100
    res2 <- 2*round(res1/2)/100 #round to nearest even integer
    strikelist[a] <- res2
  }
  
  # calculating the days to maturity of a replicating option
  
  Spot45 <- data[dim(data)[1]-(Maturity1-1), dim(data)[2]-2] / 1000 #spot price at Maturity 45 days
  date45 <- data[dim(data)[1]-(Maturity1-1), dim(data)[2]]
  
  Maturity2 <- data.frame(datanew) %>% filter(date == date45)
  Maturity2 <- Maturity2$T
  
  results <- data.frame(matrix(vector(), length(freq), length(Moneyness)))
  
  for (j in 1:length(strikelist)) { #hedging calls with different strike prices
    Set1 <- filter(data2, E == strikelist[j]) #filtering the data based on strike price
    Set2 <- filter(datanew2, E == strikelist[j])
    
    if (!empty(Set1) && !empty(Set2)) {
      
      Mat45 <- Set1 %>% filter(T == Maturity1/252) #selecting initial day 45
      S <- Mat45$S
      E <- Mat45$E
      r <- Mat45$r
      T0 <- Mat45$T
      sigma <- Mat45$volatility
      Cobs <- Mat45$Cobs
      N <- 1000
      
      MatRep <- Set2 %>% filter(T == Maturity2/252) #selecting initial day that has same date than Maturity1
      SRep <- MatRep$S
      ERep <- MatRep$E
      rRep <- MatRep$r
      T0Rep <- MatRep$T
      sigmaRep <- MatRep$volatility
      CobsRep <- MatRep$Cobs
      
      for (d in 1:(Maturity1-1)) { #check if initial rows are empty or sigma is NA
        if (empty(Mat45) || empty(MatRep)) {
          Mat45 <- Set1 %>% filter(T == (Maturity1-d)/252)
          S <- Mat45$S
          E <- Mat45$E
          r <- Mat45$r
          T0 <- Mat45$T
          sigma <- Mat45$volatility
          Cobs <- Mat45$Cobs
          
          MatRep <- Set2 %>% filter(T == (Maturity2-d)/252)
          SRep <- MatRep$S
          ERep <- MatRep$E
          rRep <- MatRep$r
          T0Rep <- MatRep$T
          sigmaRep <- MatRep$volatility
          CobsRep <- MatRep$Cobs
          
        } else if (is.na(sigma) || is.na(sigmaRep)) {
          Mat45 <- Set1 %>% filter(T == (Maturity1-d)/252)
          S <- Mat45$S
          E <- Mat45$E
          r <- Mat45$r
          T0 <- Mat45$T
          sigma <- Mat45$volatility
          Cobs <- Mat45$Cobs
          
          MatRep <- Set2 %>% filter(T == (Maturity2-d)/252)
          SRep <- MatRep$S
          ERep <- MatRep$E
          rRep <- MatRep$r
          T0Rep <- MatRep$T
          sigmaRep <- MatRep$volatility
          CobsRep <- MatRep$Cobs
          
        } else {
          break
        }
      }
      
      
      
      for (f in 1:length(freq)) { #hedging with different rehedging frequencies
        
        #initial portfolio
        
        delta45 <- EuropeanOption(type="call", underlying=S, strike=E,
                                  riskFreeRate=r, maturity=T0, volatility=sigma,
                                  dividendYield=0)$delta
        
        deltaRep <- EuropeanOption(type="call", underlying=SRep, strike=ERep,
                                   riskFreeRate=rRep, maturity=T0Rep, volatility=sigmaRep,
                                   dividendYield=0)$delta
        
        vega45 <- EuropeanOption(type="call", underlying=S, strike=E,
                                 riskFreeRate=r, maturity=T0, volatility=sigma,
                                 dividendYield=0)$vega
        
        vegaRep <- EuropeanOption(type="call", underlying=SRep, strike=ERep,
                                  riskFreeRate=rRep, maturity=T0Rep, volatility=sigmaRep,
                                  dividendYield=0)$vega
        
        NHedgeStock <- round(N * (delta45 - (vega45/vegaRep)*deltaRep)) # number of indice to hold
        
        NHedgeRep <- round(N * (vega45/vegaRep)) # amount of replicating options to hold
        
        N * Cobs - NHedgeStock * S - NHedgeRep * CobsRep # value of the initial portfolio
        
        # Dynamic hedging
        C0 <- Cobs
        S0 <- S
        C0Rep <- CobsRep
        S0Rep <- SRep
        
        Asum <- 0
        n <- 0
        k <- 0 #rehedging frequency
        
        for (time in 1:(Mat45$T*252 - 1)){
          i <- ((Mat45$T*252) - time)
          iRep <- ((MatRep$T*252) - time)
          NewMat <- Set1 %>% filter(T == i/252)
          NewMatRep <- Set2 %>% filter(T == iRep/252)
          if (!empty(NewMat) && !empty(NewMatRep)) {
            if (!is.na(NewMat$volatility) && !is.na(NewMatRep$volatility)) {
              C1 <- NewMat$Cobs
              dC <- N*(C1 - C0)
              C0 <- C1
              S1 <- NewMat$S
              dS <- NHedgeStock*(S1 - S0)
              S0 <- S1
              C1Rep <- NewMatRep$Cobs
              dCRep <- NHedgeRep*(C1Rep - C0Rep)
              C0Rep <- C1Rep
              A <- (dC - dS - dCRep)^2 
              Asum <- Asum + A
              n <- n + 1
              k <- k + 1
              if (((k %% freq[f]) == 0)) { #rehedge when remainder = 0, i.e. every other day
                Snew <- NewMat$S
                Enew <- NewMat$E
                rnew <- NewMat$r
                T0new <- NewMat$T
                sigmanew <- NewMat$volatility
                
                SnewRep <- NewMatRep$S
                EnewRep <- NewMatRep$E
                rnewRep <- NewMatRep$r
                T0newRep <- NewMatRep$T
                sigmanewRep <- NewMatRep$volatility
                
                deltanew <- EuropeanOption(type="call", underlying=Snew, strike=Enew,
                                           riskFreeRate=rnew, maturity=T0new, volatility=sigmanew,
                                           dividendYield=0)$delta
                
                deltaRep <- EuropeanOption(type="call", underlying=SnewRep, strike=EnewRep,
                                           riskFreeRate=rnewRep, maturity=T0newRep, volatility=sigmanewRep,
                                           dividendYield=0)$delta
                
                vega45 <- EuropeanOption(type="call", underlying=S, strike=E,
                                         riskFreeRate=r, maturity=T0, volatility=sigma,
                                         dividendYield=0)$vega
                
                vegaRep <- EuropeanOption(type="call", underlying=SnewRep, strike=EnewRep,
                                          riskFreeRate=rnewRep, maturity=T0newRep, volatility=sigmanewRep,
                                          dividendYield=0)$vega
                
                NHedgeStock <- round(N * (delta45 - (vega45/vegaRep)*deltaRep)) # number of stocks to hold
                
                NHedgeRep <- round(N * (vega45/vegaRep)) # amount of replicating options to hold
                
              }
            }
          }
        }
        TotalError <- Asum / n
        results[f,j] <- TotalError
      }
    }
  }
  for (i in 1:(length(Moneyness))) {
    for (j in 1:(length(freq))) {
      if (!is.na(results[j,i])) {
        monthresults[l,(i+(j-1)*length(Moneyness))] <- results[j,i]
      } 
    }
  }
  
}

averages3 <- colMeans(monthresults, na.rm = TRUE) # average of each month
sd3 <- sapply(monthresults, sd, na.rm = TRUE) # sd of each month





### 10. Delta-vega hedging with respect to moneyness and initial days to maturity. All worksheets excl. December ###


months <- c(1:11)

Moneyness <- c(0.8, 0.9, 1, 1.1) #hedging with different moneyness values
maturities <- c(60, 45, 30, 15) #hedging calls with different initial day

monthresults <- data.frame(matrix(vector(), length(months), length(Moneyness)*length(maturities)))

for (l in 1:length(months)) {
  
  data <- read_excel("isx2010C.xls", sheet = l) 
  data <- data %>% rename(T = names(data)[1]) %>% rename(S = names(data)[dim(data)[2]-2]) %>% rename(r = names(data)[dim(data)[2]-1]) #rename variables
  dates1 <- data[,dim(data)[2]]
  data <- data[,1:dim(data)[2]-1] #drop the last column (date)
  data <- mutate_all(data, function(x) as.numeric(as.character(x)))
  data["date"] <- dates1
  data <- data %>% rename(date = names(data)[dim(data)[2]])
  data <- data %>% mutate(date=as.Date(date, format = "%d.%m.%Y"))
  
  data1 <- data %>% pivot_longer(-c(T, r, S, date), names_to = "E", values_to = "Cobs" ) %>% filter(!is.na(Cobs)) %>% 
    mutate(E = map_dbl(E, as.numeric)) %>% mutate(r = r/100, T = T/252, S = S/1000, E = E/1000, Cobs = Cobs/1000) #converting to tidy data, and scaling the values
  
  scaleOversized <- function(x){ #Scaling outliers
    if(x > 1) x <- x / 1000
    return(x)
  }
  data1 <- data1 %>% mutate(Cobs = map_dbl(Cobs, scaleOversized))
  
  
  ### Calculating implied volatility ###
  
  vola <- function(Cobs, S, E, r, T){
    result <- try(EuropeanOptionImpliedVolatility(type="call", value=Cobs, underlying=S, strike=E, dividendYield=0, riskFreeRate=r, maturity=T, volatility=0.1), silent = TRUE)
    if(!is.numeric(result)){
      result <- NA
    }
    return(result)
  }
  
  data2 <- data1 %>% mutate(volatility = pmap_dbl(data1[, c('Cobs', 'S', 'E', 'r', 'T')], vola))
  
  #### reading the February data ###
  
  datanew <- read_excel("isx2010C.xls", sheet = (l+1)) 
  datanew <- datanew %>% rename(T = names(datanew)[1]) %>% rename(S = names(datanew)[dim(datanew)[2]-2]) %>% rename(r = names(datanew)[dim(datanew)[2]-1]) #rename variables
  dates2 <- datanew[,dim(datanew)[2]]
  datanew <- datanew[,1:dim(datanew)[2]-1] #drop the last column (date)
  datanew <- mutate_all(datanew, function(x) as.numeric(as.character(x)))
  datanew["date"] <- dates2
  datanew <- datanew %>% mutate(date=as.Date(date, format = "%d.%m.%Y"))
  
  datanew1 <- datanew %>% pivot_longer(-c(T, r, S, date), names_to = "E", values_to = "Cobs" ) %>% filter(!is.na(Cobs)) %>% 
    mutate(E = map_dbl(E, as.numeric)) %>% mutate(r = r/100, T = T/252, S = S/1000, E = E/1000, Cobs = Cobs/1000) #converting to tidy data, and scaling the values
  
  scaleOversized <- function(x){ #Scaling outliers
    if(x > 1) x <- x / 1000
    return(x)
  }
  datanew1 <- datanew1 %>% mutate(Cobs = map_dbl(Cobs, scaleOversized))
  
  
  ### Calculating implied volatility ###
  
  vola <- function(Cobs, S, E, r, T){
    result <- try(EuropeanOptionImpliedVolatility(type="call", value=Cobs, underlying=S, strike=E, dividendYield=0, riskFreeRate=r, maturity=T, volatility=0.1), silent = TRUE)
    if(!is.numeric(result)){
      result <- NA
    }
    return(result)
  }
  
  datanew2 <- datanew1 %>% mutate(volatility = pmap_dbl(datanew1[, c('Cobs', 'S', 'E', 'r', 'T')], vola))
  
  results <- data.frame(matrix(vector(), length(maturities), length(Moneyness)))
  
  for (j in 1:length(Moneyness)) { #hedging calls with different strike prices
    
    for (f in 1:length(maturities)) { #hedging with different rehedging frequencies
      
     
      ### Strike prices ###
      
      Maturity1 <- maturities[f]
      
      Spot45 <- data[dim(data)[1]-(Maturity1-1), dim(data)[2]-2] / 1000 #spot price at Maturity 45 days
      r45 <- data[dim(data)[1]-(Maturity1-1), dim(data)[2]-1] / 100 #r at Maturity 45 days
      
      calcstrike <- function(x, r, S, T) {
        res <- x*S*exp(r*T)
        return(res)
      }
      
      strikelist <- c()
      for (a in 1:length(Moneyness)) { #calculating corresponding strike prices from moneyness-vector
        res1 <- calcstrike(Moneyness[a], r45, Spot45, (Maturity1/252)) * 100
        res2 <- 2*round(res1/2)/100 #round to nearest even integer
        strikelist[a] <- res2
      }
      
      # calculating the days to maturity of a replicating option
      
      Spot45 <- data[dim(data)[1]-(Maturity1-1), dim(data)[2]-2] / 1000 #spot price at Maturity 45 days
      date45 <- data[dim(data)[1]-(Maturity1-1), dim(data)[2]]
      
      Maturity2 <- data.frame(datanew) %>% filter(date == date45)
      Maturity2 <- Maturity2$T
      
      Set1 <- filter(data2, E == strikelist[j]) #filtering the data based on strike price
      Set2 <- filter(datanew2, E == strikelist[j])
      
      if (!empty(Set1) && !empty(Set2)) {
        
        Mat45 <- Set1 %>% filter(T == Maturity1/252) #selecting initial day 45
        S <- Mat45$S
        E <- Mat45$E
        r <- Mat45$r
        T0 <- Mat45$T
        sigma <- Mat45$volatility
        Cobs <- Mat45$Cobs
        N <- 1000
        
        MatRep <- Set2 %>% filter(T == Maturity2/252) #selecting initial day that has same date than Maturity1
        SRep <- MatRep$S
        ERep <- MatRep$E
        rRep <- MatRep$r
        T0Rep <- MatRep$T
        sigmaRep <- MatRep$volatility
        CobsRep <- MatRep$Cobs
        
        for (d in 1:(Maturity1-1)) { #check if initial rows are empty or sigma is NA
          if (empty(Mat45) || empty(MatRep)) {
            Mat45 <- Set1 %>% filter(T == (Maturity1-d)/252)
            S <- Mat45$S
            E <- Mat45$E
            r <- Mat45$r
            T0 <- Mat45$T
            sigma <- Mat45$volatility
            Cobs <- Mat45$Cobs
            
            MatRep <- Set2 %>% filter(T == (Maturity2-d)/252)
            SRep <- MatRep$S
            ERep <- MatRep$E
            rRep <- MatRep$r
            T0Rep <- MatRep$T
            sigmaRep <- MatRep$volatility
            CobsRep <- MatRep$Cobs
            
          } else if (is.na(sigma) || is.na(sigmaRep)) {
            Mat45 <- Set1 %>% filter(T == (Maturity1-d)/252)
            S <- Mat45$S
            E <- Mat45$E
            r <- Mat45$r
            T0 <- Mat45$T
            sigma <- Mat45$volatility
            Cobs <- Mat45$Cobs
            
            MatRep <- Set2 %>% filter(T == (Maturity2-d)/252)
            SRep <- MatRep$S
            ERep <- MatRep$E
            rRep <- MatRep$r
            T0Rep <- MatRep$T
            sigmaRep <- MatRep$volatility
            CobsRep <- MatRep$Cobs
            
          } else {
            break
          }
        }
        
        #initial portfolio
        
        delta45 <- EuropeanOption(type="call", underlying=S, strike=E,
                                  riskFreeRate=r, maturity=T0, volatility=sigma,
                                  dividendYield=0)$delta
        
        deltaRep <- EuropeanOption(type="call", underlying=SRep, strike=ERep,
                                   riskFreeRate=rRep, maturity=T0Rep, volatility=sigmaRep,
                                   dividendYield=0)$delta
        
        vega45 <- EuropeanOption(type="call", underlying=S, strike=E,
                                 riskFreeRate=r, maturity=T0, volatility=sigma,
                                 dividendYield=0)$vega
        
        vegaRep <- EuropeanOption(type="call", underlying=SRep, strike=ERep,
                                  riskFreeRate=rRep, maturity=T0Rep, volatility=sigmaRep,
                                  dividendYield=0)$vega
        
        NHedgeStock <- round(N * (delta45 - (vega45/vegaRep)*deltaRep)) # number of indice to hold
        
        NHedgeRep <- round(N * (vega45/vegaRep)) # amount of replicating options to hold
        
        N * Cobs - NHedgeStock * S - NHedgeRep * CobsRep # value of the initial portfolio
        
        # Dynamic hedging
        C0 <- Cobs
        S0 <- S
        C0Rep <- CobsRep
        S0Rep <- SRep
        
        Asum <- 0
        n <- 0
        k <- 0 #rehedging frequency
        
        for (time in 1:(Mat45$T*252 - 1)){
          i <- ((Mat45$T*252) - time)
          iRep <- ((MatRep$T*252) - time)
          NewMat <- Set1 %>% filter(T == i/252)
          NewMatRep <- Set2 %>% filter(T == iRep/252)
          if (!empty(NewMat) && !empty(NewMatRep)) {
            if (!is.na(NewMat$volatility) && !is.na(NewMatRep$volatility)) {
              C1 <- NewMat$Cobs
              dC <- N*(C1 - C0)
              C0 <- C1
              S1 <- NewMat$S
              dS <- NHedgeStock*(S1 - S0)
              S0 <- S1
              C1Rep <- NewMatRep$Cobs
              dCRep <- NHedgeRep*(C1Rep - C0Rep)
              C0Rep <- C1Rep
              A <- (dC - dS - dCRep)^2 
              Asum <- Asum + A
              n <- n + 1
              k <- k + 1
              if (((k %% 2) == 0)) { #rehedge when remainder = 0, i.e. every other day
                Snew <- NewMat$S
                Enew <- NewMat$E
                rnew <- NewMat$r
                T0new <- NewMat$T
                sigmanew <- NewMat$volatility
                
                SnewRep <- NewMatRep$S
                EnewRep <- NewMatRep$E
                rnewRep <- NewMatRep$r
                T0newRep <- NewMatRep$T
                sigmanewRep <- NewMatRep$volatility
                
                deltanew <- EuropeanOption(type="call", underlying=Snew, strike=Enew,
                                           riskFreeRate=rnew, maturity=T0new, volatility=sigmanew,
                                           dividendYield=0)$delta
                
                deltaRep <- EuropeanOption(type="call", underlying=SnewRep, strike=EnewRep,
                                           riskFreeRate=rnewRep, maturity=T0newRep, volatility=sigmanewRep,
                                           dividendYield=0)$delta
                
                vega45 <- EuropeanOption(type="call", underlying=S, strike=E,
                                         riskFreeRate=r, maturity=T0, volatility=sigma,
                                         dividendYield=0)$vega
                
                vegaRep <- EuropeanOption(type="call", underlying=SnewRep, strike=EnewRep,
                                          riskFreeRate=rnewRep, maturity=T0newRep, volatility=sigmanewRep,
                                          dividendYield=0)$vega
                
                NHedgeStock <- round(N * (delta45 - (vega45/vegaRep)*deltaRep)) # number of stocks to hold
                
                NHedgeRep <- round(N * (vega45/vegaRep)) # amount of replicating options to hold
                
              }
            }
          }
        }
        TotalError <- Asum / n
        results[f,j] <- TotalError
        
      }  
    }
    
    
  }
  for (i in 1:(length(Moneyness))) {
    for (j in 1:(length(maturities))) {
      if (!is.na(results[j,i])) {
        monthresults[l,(i+(j-1)*length(Moneyness))] <- results[j,i]
      } 
    }
    
  }
}

averages4 <- colMeans(monthresults, na.rm = TRUE) # average of each month
sd4 <- sapply(monthresults, sd, na.rm = TRUE) # sd of each month


#### 11. Comparing the performance of strategies ###

deltaPerf <- mean(append(averages1, averages2[c(5:8,13:20)]))

deltavegaPerf <- mean(append(averages3, averages4[c(1:4,9:16)]))

deltaPerf / deltavegaPerf

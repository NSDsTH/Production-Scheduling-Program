import random
from deap import base, creator, tools, algorithms
from random import randint as rnd
import datetime
import time
import math
import json
import numpy as np
from SchedulePlanning import SchedulePlanning


class ProductionPlanner:
    def __init__(
        self,
        data_planning,
        population_size,
        generations,
        mutation_rate,
        n_jobs,
        prd_line,
        prd_start_date,
        man_power,
        prd_model,
        targetModel,
    ):
        self.prop_len = self.factorial(n_jobs)
        self.data_planning = data_planning
        self.population_size = (
            population_size if population_size <= self.prop_len else self.prop_len
        )
        self.generations = generations
        self.mutation_rate = mutation_rate
        self.n_jobs = n_jobs
        self.prd_line = prd_line
        self.prd_start_date = prd_start_date
        self.man_power = man_power
        self.prd_model = prd_model
        self.objective_value = []
        self.data_json = {}
        with open("parameter.json", "r") as json_file:
            self.data_json = json.load(json_file)
        self.targetModel = targetModel
        creator.create(
            "FitnessMulti", base.Fitness, weights=(-1.0, -1.0, -1.0)
        )  # Multi-objective optimization, equal weight for both objectives
        creator.create("Individual", list, fitness=creator.FitnessMulti)
        self.toolbox = base.Toolbox()
        # Create genes and individuals
        self.toolbox.register("evaluate", self.fitness_function)
        self.toolbox.register("mate", self.custom_crossover)
        self.toolbox.register("mutate", self.custom_mutation)
        self.toolbox.register("select", tools.selNSGA2)
        self.halloffame = tools.HallOfFame(maxsize=10)

    def factorial(self, n_jobs):
        if n_jobs == 1:
            return n_jobs
        else:
            return n_jobs * self.factorial(n_jobs - 1)

    @staticmethod
    def is_datetime(value):
        return isinstance(value, datetime.datetime)

    # Setting Config DEAP
    def create_individual(self, genes):
        individual = creator.Individual(genes)
        individual.fitness = creator.FitnessMulti()
        return individual

    def custom_mutation(self, individual):
        num_mutation = math.floor(len(individual) * 0.25) if len(individual) > 0 else 1
        if num_mutation <= 0 or num_mutation >= len(individual):
            print("0")
            return individual  # ไม่มีการเปลี่ยนแปลง
        positions_to_shuffle = random.sample(range(len(individual)), num_mutation)
        for _ in range(num_mutation):
            random_position = random.choice(positions_to_shuffle)
            positions_to_shuffle.remove(random_position)
            swap_position = random.choice(
                [i for i in range(len(individual)) if i not in positions_to_shuffle]
            )
            individual[random_position], individual[swap_position] = (
                individual[swap_position],
                individual[random_position],
            )
        return (individual,)

    def custom_crossover(self, ind1, ind2):
        child1, child2 = tools.cxOrdered(ind1, ind2)
        return child1, child2

    def datetime_handler(self, x):
        if isinstance(x, datetime.datetime):
            return x.timestamp()
        raise TypeError("Unknown type", x)

    def get_setup_time(self, list_model):
        count = 0
        i = 1
        while i < len(list_model):
            if list_model[i] != list_model[i - 1]:
                count += 1
            i += 1
        return count

    def fitness_function(self, chromosome):
        line = self.data_planning[chromosome[0]][-1]
        batch_chromosome = [self.data_planning[i] for i in chromosome]
        list_setup_time = [batch[1] for batch in batch_chromosome]
        setup_time = self.get_setup_time(list_setup_time)
        (
            ot_score,
            inventory_day,
            sale_date_diff,
            ot_batch,
            date_schedule,
        ) = SchedulePlanning(self.prd_start_date).simulate_plan(
            batch_chromosome, self.man_power[line], self.prd_model, self.targetModel
        )
        return (ot_score, setup_time, (inventory_day / 30))

    def create_initial_chromosome(self):
        initial_chromosome = list(range(len(self.data_planning)))
        random.shuffle(initial_chromosome)
        return initial_chromosome

    def create_initial_population(self):
        population = []
        while len(population) < self.population_size:
            chromosome = self.create_initial_chromosome()
            if chromosome not in population:
                population.append(chromosome)
        return population

    @staticmethod
    def crossover(parent1, parent2):
        crossover_point = random.randint(0, len(parent1))
        child1 = parent1[:crossover_point] + [
            gene for gene in parent2 if gene not in parent1[:crossover_point]
        ]
        child2 = parent2[:crossover_point] + [
            gene for gene in parent1 if gene not in parent2[:crossover_point]
        ]
        return child1, child2

    def genetic_algorithm(self):
        population = self.create_initial_population()
        population = [self.create_individual(item) for item in population]
        n_jobs = len(population)
        # Run the NSGA-II algorithm
        algorithms.eaMuPlusLambda(
            population,
            self.toolbox,
            mu=n_jobs,
            lambda_=n_jobs,
            cxpb=0.4,
            mutpb=0.6,
            ngen=self.generations,
            stats=None,
            # halloffame=self.halloffame,
            verbose=False,
        )
        # Get the best individuals from the final population
        best_inds = tools.sortNondominated(
            population, len(population), first_front_only=True
        )[0]
        # print("Best individuals:")
        # for ind in best_inds:
        #     print("Individual:", ind)
        #     print("Objective Values:", ind.fitness.values)
        return (
            best_inds[0],
            best_inds[0].fitness.values[0]
            + best_inds[0].fitness.values[1]
            + best_inds[0].fitness.values[2],
        )

    def run(self):
        if len(self.data_planning) < 4:
            (
                ot_score,
                inventory_day,
                sale_date_diff,
                ot_batch,
                date_schedule,
            ) = SchedulePlanning(self.prd_start_date).simulate_plan(
                self.data_planning,
                self.man_power[self.data_planning[0][-1]],
                self.prd_model,
                self.targetModel,
            )
            print("Less than 5")
            return self.data_planning, 0, date_schedule
        else:
            prd_schedule = []
            best_order, best_fitness = self.genetic_algorithm()
            for index in best_order:
                prd_schedule.append(self.data_planning[index])
            (
                ot_score,
                inventory_day,
                sale_date_diff,
                ot_batch,
                date_schedule,
            ) = SchedulePlanning(self.prd_start_date).simulate_plan(
                prd_schedule,
                self.man_power[prd_schedule[0][-1]],
                self.prd_model,
                self.targetModel,
            )
            return prd_schedule, best_fitness, date_schedule
